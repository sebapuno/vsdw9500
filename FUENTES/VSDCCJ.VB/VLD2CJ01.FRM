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
Dim GetAis As String * 70
Dim Correlativo As String
Dim Respuesta As Integer
Dim BaseVisado As Database
Dim DatoProducto As String * 11
Dim PunteroVsdPrd  As Recordset
Dim PunteroTabPrd  As Recordset
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    PunteroTabPrd.Index = "Tab_Prd"
    Respuesta = AisGetUsr(GetAis)
    IdUsuario = Format(Mid(GetAis, 1, 8), "00000000")
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", Producto
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    DatoProducto = Mid$(PunteroVsdPrd("Prd_Data"), 61, 11)
    If ExisteCuenta(DatoProducto, PunteroTabPrd("Prd_Vista_Dst")) Then
      Respuesta = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
    End If
    PunteroTabPrd.Close
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub
Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  Else
    PoneValorINI = False
  End If
End Function
Private Function ExisteCuenta(ByVal Llave As String, ByVal Vista As String) As Boolean
Dim Usuario As String
Dim NodoDestino As String * 4
Dim AplicacionDestino As String * 4
  Modo = "L"
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  ExisteCuenta = False
  If Trim$(NodoDestino) <> "" Then
    Usuario = IdUsuario + Vista + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave, "1K", True, False)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      If (Mid$(CadenaDatos, 70, 8) = "00000000") And (InStr("03,05", Mid$(CadenaDatos, 23, 2)) <> 0) Then
        MsgBox "La Cuenta Juvenil Asociada" + Chr(10) + _
                      "al Producto Ya Está Cerrada", 0 + 16, "Problemas de Validación"
      Else
        ExisteCuenta = True
      End If
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
        MsgBox "No Existe la Cuenta Corriente" + Chr(10) + _
                      "Juvenil Asociada al Producto", 0 + 16, "Problemas de Validación"
      End If
    End If
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
Dim BaseEnvioSRM As Database
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  PunteroVistas.Index = PunteroVistas.Indexes(0).Name
  PunteroVistas.Seek "=", Vista
  If PunteroVistas.NoMatch Then
    DeterminaNodo = ""
  Else
    If Environ$("DESA") = "DESA" Then
      DeterminaNodo = PunteroVistas("Vta_NodoDesa")
    Else
      DeterminaNodo = PunteroVistas("Vta_NodoProd")
    End If
    Aplicacion = PunteroVistas("Vta_Aplicacion")
  End If
  PunteroVistas.Close
  BaseEnvioSRM.Close
End Function
