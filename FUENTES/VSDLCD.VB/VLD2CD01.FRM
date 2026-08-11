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
Dim IdUsuario   As String * 8
Dim CadenaDatos As String
Dim Data_IBM    As String
Dim BaseVisado  As Database

Function Actualiza(ByVal Data_Visa As String)
   Dim Data_Ret As String
   Dim Fech_Tmp As String * 8
   
   Data_Ret = Mid$(Data_IBM, 1, 50)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 51, 14)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 65, 14)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 79, 14)

   'la idea es que el campo LCD_FEC_RENOVACION no sea modificado
   If Mid$(Data_IBM, 227, 8) = Mid$(Data_Visa, 227, 8) Then
      Data_Ret = Data_Ret + Mid$(Data_IBM, 93, 8)
   Else
      Data_Ret = Data_Ret + Format$(Date, "yyyymmdd")
   End If

   Data_Ret = Data_Ret + Mid$(Data_IBM, 101, 96)

   'la idea es que el campo LCD_FEC_APERTURA no sea modificado
   Data_Ret = Data_Ret + Mid$(Data_IBM, 197, 8)

   Data_Ret = Data_Ret + Mid$(Data_IBM, 205, 13)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 218, 1)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 219, 3)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 222, 15)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 237, 1)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 238, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 242, 6)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 248, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 252, 8)
   Data_Ret = Data_Ret + Mid$(Data_Visa, 260, 4)
   Data_Ret = Data_Ret + Mid$(Data_IBM, 264, 7)

   If Mid$(Data_IBM, 227, 8) = Mid$(Data_Visa, 227, 8) Then
      ' MODIFICACION
      Data_Ret = Data_Ret + Mid$(Data_IBM, 271, 8)
   Else
      ' RENOVACION
      Data_Ret = Data_Ret + Format$(Date, "yyyymmdd")
   End If
   
   'Se guarda la fecha del campo LCD_FEC_ULT_ACTUALIZ.
   'la idea es guardar dicha fecha en el campo LCD_FEC_PENUL_ACTUAL
   Fech_Tmp = Mid$(Data_IBM, 279, 8)

   Data_Ret = Data_Ret + Format$(Date, "yyyymmdd")

   'Se guarda la fecha proveniente del campo LCD_FEC_ULT_ACTUALIZ
   'en el campo LCD_FEC_PENUL_ACTUAL
   Data_Ret = Data_Ret + Fech_Tmp
   
   Data_Ret = Data_Ret + Mid$(Data_IBM, 295, 21)

   Actualiza = Data_Ret

End Function


Private Sub Form_Load()
Dim Reg            As String
Dim Dat            As String
Dim Data_Visa      As String
Dim Operacion      As String
Dim Correlativo    As String
Dim LineaDeCredito As String
Dim PunteroVsdPrd  As Recordset

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    ' FI2026 - era: PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
    ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
    ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
    ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
    PunteroVsdPrd.Index = BaseVisado.TableDefs("Vsd_Prd").Indexes(0).Name
    IdUsuario = ObtieneRutUsuarioLogueado()
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
    Operacion = Mid$(PunteroVsdPrd("Prd_Data"), 40, 1)
    LineaDeCredito = Mid$(PunteroVsdPrd("Prd_Data"), 10, 11)
    Reg = Mid$(PunteroVsdPrd("Prd_Data"), 1, 60)
    Dat = Mid$(PunteroVsdPrd("Prd_Data"), 61)
    If Operacion = "M" Or Operacion = "R" Then
      If ExisteLineaDeCredito(LineaDeCredito, "0392", True) Then
         Data_Visa = Actualiza(Dat)
         PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
         PunteroVsdPrd.Edit
         PunteroVsdPrd("Prd_Data") = Reg + Data_Visa
         PunteroVsdPrd.Update
         If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
           MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                  "Error de Escritura"
         End If
      End If
    Else
      If Not ExisteLineaDeCredito(LineaDeCredito, "0392", False) Then
         If ExisteCuentaCorriente(LineaDeCredito) Then
            If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
               MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                      "Error de Escritura"
            End If
         End If
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
Private Function ExisteLineaDeCredito(ByVal Llave As String, ByVal Vista As String, ByVal Estado As Boolean) As Boolean
Dim DataTmp As String

  Modo = "V0000"
  ExisteLineaDeCredito = ObtieneInformacion(Vista, Llave, False, Estado)
  If ExisteLineaDeCredito Then
     ' Este codigo transforma de 0DDMMSSAA a SSAAMMDD (año 2000)
     DataTmp = Mid$(CadenaDatos, 1, 26) + Mid$(CadenaDatos, 26 + 1, 42)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 48, 4) + Mid$(CadenaDatos, 26 + 46, 2) + Mid$(CadenaDatos, 26 + 44, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 52, 42)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 98, 4) + Mid$(CadenaDatos, 26 + 96, 2) + Mid$(CadenaDatos, 26 + 94, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 102, 9)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 116, 4) + Mid$(CadenaDatos, 26 + 114, 2) + Mid$(CadenaDatos, 26 + 112, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 120, 14)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 139, 4) + Mid$(CadenaDatos, 26 + 137, 2) + Mid$(CadenaDatos, 26 + 135, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 143, 49)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 195, 4) + Mid$(CadenaDatos, 26 + 193, 2) + "01"
     
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 203, 4) + Mid$(CadenaDatos, 26 + 201, 2) + Mid$(CadenaDatos, 26 + 199, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 207, 22)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 233, 4) + Mid$(CadenaDatos, 26 + 231, 2) + Mid$(CadenaDatos, 26 + 229, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 237, 36)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 277, 4) + Mid$(CadenaDatos, 26 + 275, 2) + Mid$(CadenaDatos, 26 + 273, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 285, 4) + Mid$(CadenaDatos, 26 + 283, 2) + Mid$(CadenaDatos, 26 + 281, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 293, 4) + Mid$(CadenaDatos, 26 + 291, 2) + Mid$(CadenaDatos, 26 + 289, 2)
     DataTmp = DataTmp + Mid$(CadenaDatos, 26 + 297, 21)
     CadenaDatos = DataTmp
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
            ' En año 2000 SD20 cambiar 170 a 176 listo!!
            If Mid$(CadenaDatos, 15 + 176, 1) <> " " And Mid$(CadenaDatos, 15 + 176, 1) <> "0" Then
              MsgBox "La Línea de Crédito Está Cerrada", 0 + 16, "Problemas de Validación"
            Else
              ' En año 2000 SD20 cambiar 173 a 179 listo!!
              If Mid$(CadenaDatos, 15 + 179, 1) = "4" Then
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

Private Function ExisteCuentaCorriente(ByVal Llave As String) As Boolean
  Modo = "L"
  Vista = "0391"
  ExisteCuentaCorriente = ObtieneInformacion(Vista, "0" + Mid$(Llave, 2, 8) + Format((Val(Mid$(Llave, 10, 2)) + 10) Mod 11, "00"), True, False)
End Function
Private Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim BaseEnvioSRM  As Database
Dim PunteroVistas As Recordset
Dim PunteroPrd    As Recordset

  DeterminaNodo = SINNODO
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - era: PunteroVistas.Index = PunteroVistas.Indexes(0).Name
  ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
  ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
  ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
  PunteroVistas.Index = BaseEnvioSRM.TableDefs("Vistas").Indexes(0).Name
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
  
    ' Analiza si el producto ya es 2000
    'Set PunteroPrd = BaseVisado.OpenRecordset("TAB_PRD", dbOpenTable)
    ' FI2026 - 'PunteroPrd.Index = 'PunteroPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    'PunteroPrd.Seek "=", Producto
    'If PunteroPrd("Prd_2000") And PunteroVistas("Vta_Aplicacion") = "SPLA" Then
    '   Aplicacion = "SPL2"
    'Else
       Aplicacion = PunteroVistas("Vta_Aplicacion")
    'End If
    'PunteroPrd.Close
  
  End If
  PunteroVistas.Close
  BaseEnvioSRM.Close
End Function
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function
