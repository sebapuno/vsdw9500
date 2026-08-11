VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1035
   ClientTop       =   990
   ClientWidth     =   4350
   Height          =   2220
   Left            =   975
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4350
   Top             =   645
   Visible         =   0   'False
   Width           =   4470
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario   As String * 8
Dim CadenaDatos As String
Dim Data_Vax    As String
Dim BaseVisado  As Database

Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Data_Visa      As String
Dim DataCampana    As String
Dim GestionLCD     As Boolean
Dim Operacion      As String
Dim Correlativo    As String
Dim LineaDeCredito As String * 11
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    ' FI2026 - era: PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
    ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
    ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
    PunteroVsdPrd.Index = BaseVisado.TableDefs("Vsd_Prd").Indexes(0).Name
    PunteroTabPrd.Index = "Tab_Prd"
    GestionLCD = False
    IdUsuario = ObtieneRutUsuarioLogueado()
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", Producto
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    Operacion = Mid$(PunteroVsdPrd("Prd_Data"), 40, 1)
    LineaDeCredito = Mid$(PunteroVsdPrd("Prd_Data"), 11, 11)
    Reg = Mid$(PunteroVsdPrd("Prd_Data"), 1, 60)
    Dat = Mid$(PunteroVsdPrd("Prd_Data"), 61)
    If ExisteLineaDeCredito(LineaDeCredito, PunteroTabPrd("Prd_Vista_Dst"), True) Then
       'If Mid$(Dat, 171, 8) <> Mid$(Data_Vax, 171, 8) Then
       '   DataCampana = Es_FCampana(Mid(Dat, 4, 9), LineaDeCredito)
       '   If Mid$(DataCampana, 1, 3) <> "SRM" Then
       '      If Mid$(DataCampana, 1, 3) <> "XXX" Then
       '         GestionLCD = True
       '         DataCampana = DataCampana + Format$(Val(Mid$(Dat, 139, 11)) - Val(Mid$(Data_Vax, 139, 11)), "0000000000000")
                'PunteroVsdPrd.Seek "=", Rut, "GLC", Val(Correlativo)
                'If PunteroVsdPrd.NoMatch Then
                '   PunteroVsdPrd.AddNew
                '   PunteroVsdPrd("Prd_RutCli") = Rut
                '   PunteroVsdPrd("Prd_CodProd") = "GLC"
                '   PunteroVsdPrd("Prd_Correl") = Val(Correlativo)
                '   PunteroVsdPrd("Prd_Data") = Reg + DataCampana
                '   PunteroVsdPrd.Update
                'End If
        '     End If
        '  Else
        '     MsgBox "No Se Puede Determinar Gestión de LCD", 0 + 48, "Error de Validacion"
        '     End
        '  End If
        'End If
       Data_Visa = Actualiza(Dat, GestionLCD)
       PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
       PunteroVsdPrd.Edit
       PunteroVsdPrd("Prd_Data") = Reg + Data_Visa
       PunteroVsdPrd.Update
    
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
Function Actualiza(ByVal Data_Visa As String, ByVal GestionLCD As Boolean)
   Dim Data_Ret As String
   
   Data_Ret = Mid$(Data_Vax, 1, 28)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 29, 2)
   Data_Ret = Data_Ret + Mid$(Data_Vax, 31, 113)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 144, 46)
   Data_Ret = Data_Ret + Mid$(Data_Vax, 190, 133)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 323, 19)
   
   Actualiza = Data_Ret

End Function

Private Function Es_FCampana(ByVal Rut1 As String, ByVal Cuenta1 As String) As String
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4

  NodoDestino = DeterminaNodo(AplicacionDestino, "2147")
  Modo = "L"
  Llave = Rut1 + "0" + Cuenta1
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + "2147" + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
       Es_FCampana = Mid$(CadenaDatos, 7, 26)
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
         Es_FCampana = "XXX"
      Else
         Es_FCampana = "SRM"
      End If
    End If
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
    Es_FCampana = "SRM"
  End If
End Function

Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String, ByVal Estado As Boolean) As Boolean
Dim VarTmp As String
  Modo = "V"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, "0" + Llave, False, Estado)
  If ExisteLineaDeCredito Then
  
     Data_Vax = Mid$(CadenaDatos, 28)
     
     ' SACAR CUANDO LA VISTA 4391 ESTE 2000
     'VarTmp = Mid$(Data_Vax, 1, 42)
     
     'If Mid$(Data_Vax, 43, 6) = "000000" Then
     '   VarTmp = VarTmp + "00" + Mid$(Data_Vax, 43, 64)
     'Else
     '   VarTmp = VarTmp + "19" + Mid$(Data_Vax, 43, 64)
     'End If
     'If Mid$(Data_Vax, 107, 6) = "000000" Then
     '   VarTmp = VarTmp + "00" + Mid$(Data_Vax, 107, 15)
     'Else
     '   VarTmp = VarTmp + "19" + Mid$(Data_Vax, 107, 15)
     'End If
     'If Mid$(Data_Vax, 122, 6) = "000000" Then
     '   VarTmp = VarTmp + "00" + Mid$(Data_Vax, 122, 35)
     'Else
     '   VarTmp = VarTmp + "19" + Mid$(Data_Vax, 122, 35)
     'End If
     'If Mid$(Data_Vax, 157, 6) = "000000" Then
     '   VarTmp = VarTmp + "00" + Mid$(Data_Vax, 157, 6)
     'Else
     '   VarTmp = VarTmp + "19" + Mid$(Data_Vax, 157, 6)
     'End If
     'If Mid$(Data_Vax, 163, 6) = "999999" Then
     '   VarTmp = VarTmp + "99" + Mid$(Data_Vax, 163, 150) '999999
     'Else
     '   VarTmp = VarTmp + "19" + Mid$(Data_Vax, 163, 150) '999999
     'End If
     'If Mid$(Data_Vax, 313, 6) = "000000" Then
     '   VarTmp = VarTmp + "00" + Mid$(Data_Vax, 313, 18)
     'Else
     '   VarTmp = VarTmp + "19" + Mid$(Data_Vax, 313, 18)
     'End If
     
     'Data_Vax = VarTmp
     
  End If
End Function
Private Function ObtieneInformacion(ByVal Vista As String, ByVal Llave As String, ByVal SeEstaBuscandoCuenta As Boolean, ByVal DebeExistirLinea As Boolean) As Boolean
Dim Usuario As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
  ObtieneInformacion = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    Usuario = IdUsuario + Vista + Modo
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      If SeEstaBuscandoCuenta Then
        If Mid$(Llave, 3, 8) = Mid$(CadenaDatos, 18, 8) Then
          ObtieneInformacion = True
        Else
          MsgBox "No Existe la Cuenta Corriente Asociada", 0 + 16, "Problemas de Validación"
        End If
      Else
        If DebeExistirLinea Then
          If Mid$(Llave, 1, 10) = Mid$(CadenaDatos, 16, 10) Then
            If Val(Mid$(CadenaDatos, 149, 6)) <> 0 Then
              MsgBox "La Línea de Crédito Está Cerrada", 0 + 16, "Problemas de Validación"
            Else
              If Mid$(CadenaDatos, 291, 1) = "4" Then
                MsgBox "La Línea de Crédito Está" + Chr(10) + "en Proceso de ser Cerrada", 0 + 16, _
                              "Problemas de Validación"
              Else
                ObtieneInformacion = True
              End If
            End If
          Else
            MsgBox "No Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
          End If
        Else
          If Llave = Mid$(CadenaDatos, 16, 12) Then
            ObtieneInformacion = True
            'MsgBox "Ya Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
          End If
        End If
      End If
    Else
      If Mid$(CadenaDatos, 1, 6) = "000096" Then
        If SeEstaBuscandoCuenta Then
          MsgBox "No Existe la Cuenta Corriente Asociada", 0 + 16, "Problemas de Validación"
        Else
          If DebeExistirLinea Then
            MsgBox "No Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
          End If
        End If
      End If
    End If
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
End Function
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "V"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "00" + Llave + "00", True, False)
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset

     Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     ' FI2026 - era: PunteroVista.Index = PunteroVista.Indexes(0).Name
     ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
     ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
     ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
     PunteroVista.Index = BaseEnvioSRM.TableDefs("Vistas").Indexes(0).Name
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
       'Set PunteroPrd = BaseVisado.OpenRecordset("TAB_PRD", dbOpenTable)
       ' FI2026 - 'PunteroPrd.Index = 'PunteroPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
       'PunteroPrd.Seek "=", Producto
       'If PunteroPrd("Prd_2000") And PunteroVista("Vta_Aplicacion") = "SPLA" Then
       '   Aplicacion = "SPL2"
       'Else
          Aplicacion = PunteroVista("Vta_Aplicacion")
       'End If
       'PunteroPrd.Close
     End If
     PunteroVista.Close
     BaseEnvioSRM.Close
End Function

Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
