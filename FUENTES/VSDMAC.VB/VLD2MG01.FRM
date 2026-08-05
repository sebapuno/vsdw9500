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
Dim BaseVisado   As Database
Dim BaseEnvioSRM As Database


Private Sub Form_Load()
Dim Data           As String
Dim Correlativo    As String
Dim Nro_Grupo      As String
Dim Nro_Correl     As String
Dim PunteroVsdPrd  As Recordset
Dim NodoDestino    As String
Dim AplicDestino   As String
Dim Encabezado     As String

If Not App.PrevInstance Then
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
  
  IdUsuario = ObtieneRutUsuarioLogueado()
  Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
  PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
  
  If Not PunteroVsdPrd.NoMatch Then
    
    NodoDestino = DeterminaNodo(AplicDestino, "SY44")
    Nro_Grupo = Mid$(PunteroVsdPrd("Prd_Data"), 61, 10)
    Nro_Correl = Mid$(PunteroVsdPrd("Prd_Data"), 71, 3)
    
    Encabezado = "b12256781234W00000010plt_corp  exec sp_macg_consulta_estado "
    Data = Nro_Grupo + "," + Nro_Correl
    
    Ret = Llamada(NodoDestino, AplicDestino, Encabezado, Data, "08")
    
    If Mid$(Ret, 15, 1) = "V" Then
      If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
       MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
              "Error de Escritura"
      End If
    Else
       MsgBox "", 16
    End If
    
  End If
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


Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset

     Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     PunteroVista.Index = PunteroVista.Indexes(0).Name
     PunteroVista.Seek "=", Vista
     If PunteroVista.NoMatch Then
       DeterminaNodo = SINNODO
     Else
       If Environ$("DESA") = "DESA" Then
         DeterminaNodo = PunteroVista("Vta_NodoDesa")
       Else
         If PunteroVista("Vta_NodoProd") = "U???" Then
           DeterminaNodo = LCase$(Environ$("SRM_NS"))
         Else
           DeterminaNodo = PunteroVista("Vta_NodoProd")
         End If
       End If
       
       ' Analiza si el producto ya es 2000
       Set PunteroPrd = BaseVisado.OpenRecordset("TAB_PRD", dbOpenTable)
       PunteroPrd.Index = PunteroPrd.Indexes(0).Name
       PunteroPrd.Seek "=", Producto
       Aplicacion = PunteroVista("Vta_Aplicacion")
       PunteroPrd.Close
     End If
     PunteroVista.Close
     BaseEnvioSRM.Close
End Function

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
