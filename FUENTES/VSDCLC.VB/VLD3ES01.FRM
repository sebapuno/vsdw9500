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
Dim Correlativo As String
Dim BaseVisado As Database
Dim LineaDeCredito As String
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    PunteroTabPrd.Index = "Tab_Prd"
    IdUsuario = ObtieneRutUsuarioLogueado()
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", Producto
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    LineaDeCredito = Mid$(PunteroVsdPrd("Prd_Data"), 11, 11)
    If ExisteLineaDeCredito(LineaDeCredito, PunteroTabPrd("Prd_Vista_Dst")) Then
      If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
          MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                        "Error de Escritura"
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
Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, "0" + Llave, "LÃƒÆ’Ã‚Â­nea de CrÃƒÆ’Ã‚Â©dito", 16, 10)
End Function
Private Function ObtieneInformacion(ByVal Vista As String, ByVal Llave As String, ByVal DatoBuscado As String, ByVal Posicion As Integer, ByVal Largo As Integer) As Boolean
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
  ObtieneInformacion = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + Vista + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      If Mid$(Llave, 1, Largo) = Mid$(CadenaDatos, Posicion, Largo) Then
        If Val(Mid(CadenaDatos, 149, 6)) <> 0 Then
          MsgBox "La " + DatoBuscado + " EstÃƒÆ’Ã‚Â¡ Cerrada", 0 + 16, "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
        Else
          If Mid$(CadenaDatos, 291, 1) <> "0" Then
            MsgBox "La " + DatoBuscado + " EstÃƒÆ’Ã‚Â¡" + Chr(10) + "en Proceso de ser Cerrada", 0 + 16, _
                          "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
          Else
            ObtieneInformacion = True
          End If
        End If
      Else
        MsgBox "No Existe la " + DatoBuscado, 0 + 16, "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
      End If
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
        MsgBox "No Existe la " + DatoBuscado, 0 + 16, "Problemas de ValidaciÃƒÆ’Ã‚Â³n"
      End If
    End If
  Else
    MsgBox "Vista NÃƒâ€šÃ‚Âº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
Dim BaseEnvioSRM As Database
  DeterminaNodo = SINNODO
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - PunteroVistas.Index = PunteroVistas.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
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
  ctrlClave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
