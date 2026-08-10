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
Dim Data_IBM    As String

Dim BaseVisado   As Database
Dim BaseEnvioSRM As Database

Function Actualiza(ByVal Data_Visa As String, ByVal GestionLCD As Boolean)
   Dim Data_Ret As String
   Dim Fech_Tmp As String * 8
   
   Data_Ret = Mid$(Data_IBM, 1, 51)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 52, 14)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 66, 14)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 80, 14)

   'la idea es que el campo LCD_FEC_RENOVACION no sea modificado
   If Mid$(Data_IBM, 229, 8) = Mid$(Data_Visa, 229, 8) Then
      Data_Ret = Data_Ret + Mid$(Data_IBM, 94, 8)
   Else
      Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")
   End If

   Data_Ret = Data_Ret + Mid$(Data_IBM, 102, 97)

   'la idea es que el campo LCD_FEC_APERTURA no sea modificado
   Data_Ret = Data_Ret + Mid$(Data_IBM, 199, 8)

   Data_Ret = Data_Ret + Mid$(Data_IBM, 207, 13)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 220, 1)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 221, 3)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 224, 15)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 239, 1)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 240, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 244, 6)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 250, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 254, 8)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 262, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 266, 7)

   If Mid$(Data_IBM, 229, 8) = Mid$(Data_Visa, 229, 8) Or GestionLCD Then
      ' MODIFICACION
      Data_Ret = Data_Ret + Mid$(Data_IBM, 273, 8)
   Else
      ' RENOVACION
      Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")
   End If
   
   'Se guarda la fecha del campo LCD_FEC_ULT_ACTUALIZ.
   'la idea es guardar dicha fecha en el campo LCD_FEC_PENUL_ACTUAL
   Fech_Tmp = Mid$(Data_IBM, 281, 8)

   Data_Ret = Data_Ret + Format$(Date, "ddmmyyyy")

   'Se guarda la fecha proveniente del campo LCD_FEC_ULT_ACTUALIZ
   'en el campo LCD_FEC_PENUL_ACTUAL
   Data_Ret = Data_Ret + Fech_Tmp
   
   If Mid$(Data_IBM, 229, 8) <> Mid$(Data_Visa, 229, 8) And GestionLCD Then
      Data_Ret = Data_Ret + Format$(Date, "yyyy")
      Data_Ret = Data_Ret + Mid$(Data_IBM, 301, 17)
   Else
      Data_Ret = Data_Ret + Mid$(Data_IBM, 297, 21)
   End If

   Actualiza = Data_Ret

End Function


Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Data_Visa      As String
Dim DataCampana    As String
Dim GestionLCD     As Boolean
Dim Operacion      As String
Dim Correlativo    As String
Dim LineaDeCredito As String
Dim PunteroTabPrd  As Recordset
Dim PunteroVsdPrd  As Recordset

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    PunteroTabPrd.Index = "Tab_Prd"
    GestionLCD = False
    IdUsuario = ObtieneRutUsuarioLogueado()
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", Producto
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    Operacion = Mid$(PunteroVsdPrd("Prd_Data"), 40, 1)
    LineaDeCredito = Mid$(PunteroVsdPrd("Prd_Data"), 10, 11)
    Reg = Mid$(PunteroVsdPrd("Prd_Data"), 1, 60)
    Dat = Mid$(PunteroVsdPrd("Prd_Data"), 61)
    If Operacion = "M" Then
      If ExisteLineaDeCredito(LineaDeCredito, PunteroTabPrd("Prd_Vista_Dst"), True) Then
         If Mid$(Dat, 229, 8) <> Mid$(Data_IBM, 229, 8) Then
            DataCampana = Es_FCampana(Mid(Dat, 174, 9), LineaDeCredito)
            If Mid$(DataCampana, 1, 3) <> "SRM" Then
               If Mid$(DataCampana, 1, 3) <> "XXX" Then
                  GestionLCD = True
                  DataCampana = DataCampana + Format$(Val(Mid$(Dat, 53, 13)) - Val(Mid$(Data_IBM, 53, 13)), "0000000000000")
                  PunteroVsdPrd.Seek "=", Rut, "GLC", Val(Correlativo)
                  If PunteroVsdPrd.NoMatch Then
                     PunteroVsdPrd.AddNew
                     PunteroVsdPrd("Prd_RutCli") = Rut
                     PunteroVsdPrd("Prd_CodProd") = "GLC"
                     PunteroVsdPrd("Prd_Correl") = Val(Correlativo)
                     PunteroVsdPrd("Prd_Data") = Reg + DataCampana
                     PunteroVsdPrd.Update
                  End If
               End If
            Else
               MsgBox "No Se Puede Determinar Gestión de LCD", 0 + 48, "Error de Validacion"
               End
            End If
         End If
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
Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String, ByVal Estado As Boolean) As Boolean
Dim DataTmp As String

  Modo = "V0000"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, Llave, False, Estado)
  If ExisteLineaDeCredito Then
     ' nuevo Sacar una vez sea SD20 (Este codigo transforma a año 2000)
     'DataTmp = Mid$(CadenaDatos, 1, 26) + Mid$(CadenaDatos, 26 + 1, 42)
     'DataTmp = DataTmp + "19" + Mid$(CadenaDatos, 26 + 48, 2) + Mid$(CadenaDatos, 26 + 46, 2) + Mid$(CadenaDatos, 26 + 44, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 50, 42)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 96, 4) + Mid$(CadenaDatos, 26 + 94, 2) + Mid$(CadenaDatos, 26 + 92, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 100, 9)
     'DataTmp = DataTmp + "19" + Mid$(CadenaDatos, 26 + 114, 2) + Mid$(CadenaDatos, 26 + 112, 2) + Mid$(CadenaDatos, 26 + 110, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 116, 14)
     'DataTmp = DataTmp + "19" + Mid$(CadenaDatos, 26 + 135, 2) + Mid$(CadenaDatos, 26 + 133, 2) + Mid$(CadenaDatos, 26 + 131, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 137, 49)
     'DataTmp = DataTmp + "19" + Mid$(CadenaDatos, 26 + 189, 2) + Mid$(CadenaDatos, 26 + 187, 2) + "01"
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 195, 4) + Mid$(CadenaDatos, 26 + 193, 2) + Mid$(CadenaDatos, 26 + 191, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 199, 22)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 225, 4) + Mid$(CadenaDatos, 26 + 223, 2) + Mid$(CadenaDatos, 26 + 221, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 229, 36)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 269, 4) + Mid$(CadenaDatos, 26 + 267, 2) + Mid$(CadenaDatos, 26 + 265, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 277, 4) + Mid$(CadenaDatos, 26 + 275, 2) + Mid$(CadenaDatos, 26 + 273, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 285, 4) + Mid$(CadenaDatos, 26 + 283, 2) + Mid$(CadenaDatos, 26 + 281, 2)
     'DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 289, 21)
     'CadenaDatos = DataTmp
     ' hasta aqui
     Data_IBM = Mid$(CadenaDatos, 27)
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
        ObtieneInformacion = True
      Else
        If DebeExistirLinea Then
          If Llave = Mid$(CadenaDatos, 15 + 1, 11) Then
            ' En año 2000 SD20 cambiar 170 a 173
            If Mid$(CadenaDatos, 15 + 176, 1) <> " " And Mid$(CadenaDatos, 15 + 176, 1) <> "0" Then
              MsgBox "La Línea de Crédito Está Cerrada", 0 + 16, "Problemas de Validación"
            Else
              ' En año 2000 SD20 cambiar 173 a 176
              If Mid$(CadenaDatos, 15 + 180, 1) = "4" Then
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
          If Llave = Mid$(CadenaDatos, 16, 11) Then
            ObtieneInformacion = True
            MsgBox "Ya Existe la Línea de Crédito", 0 + 16, "Problemas de Validación"
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
Private Function ExisteCuentaCorriente(ByVal Llave As String, ByVal Vista As String) As Boolean
  Modo = "L"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "0" + Mid$(Llave, 2, 8) + Format((Val(Mid$(Llave, 11, 2)) + 10) Mod 11, "00"), True, False)
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset

     Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     ' FI2026 - PunteroVista.Index = PunteroVista.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
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
       ' FI2026 - PunteroPrd.Index = PunteroPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
       PunteroPrd.Seek "=", Producto
       Aplicacion = PunteroVista("Vta_Aplicacion")
       PunteroPrd.Close
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
