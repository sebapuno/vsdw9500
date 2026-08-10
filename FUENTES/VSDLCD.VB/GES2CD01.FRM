VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1545
   ClientTop       =   1725
   ClientWidth     =   3630
   Height          =   2220
   Left            =   1485
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3630
   Top             =   1380
   Visible         =   0   'False
   Width           =   3750
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario As String * 8
Dim CadenaDatos As String

Private Sub Form_Load()
Dim GetAis        As String * 70
Dim DatoRut       As String * 9
Dim Correlativo   As String
Dim Respuesta     As Integer
Dim BaseVisado    As Database
Dim BaseEnvioSRM  As Database
Dim DatoProducto  As String * 11
Dim PunteroVsdPrd As Recordset
Dim PunteroTabPrd As Recordset
Dim PunteroTxSat  As Recordset

Dim LlaveSAT    As String
Dim DataSAT     As String
Dim Nro_Linea   As String
Dim Rut_Clte    As String
Dim DataC       As String
Dim Oficina     As String * 3
Dim FechaAAMMDD As String
Dim TipoVisado  As String * 2

  If InStr("00,54,55,56", Environ$("NODOM")) = 0 Then
     TipoVisado = "01" 'regional
  Else
     TipoVisado = "02" 'metro
  End If
  Oficina = Environ$("OFICINA")
  D_TRANSAC = "0011"
  
  If Not App.PrevInstance Then
    Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
    Set PunteroTxSat = BaseEnvioSRM.OpenRecordset("SATGESTION", dbOpenTable)
    
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    PunteroTabPrd.Index = "Tab_Prd"
    ' FI2026 SHG - AisGetUsr (WAIS.DLL 16-bit) deshabilitado
    'Respuesta = AisGetUsr(GetAis)
    IdUsuario = Format(Mid(GetAis, 1, 8), "00000000")
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", "GLC"
    PunteroVsdPrd.Seek "=", Rut, "GLC", Val(Correlativo)
    If Not PunteroVsdPrd.EOF Then
       If Not PunteroVsdPrd.NoMatch Then
          DataC = Mid$(PunteroVsdPrd("Prd_Data"), 61)
          LlaveSAT = IdUsuario + D_TRANSAC + "G" + "0000" + "GLCD" + "Gestion AumCupo" + "99"

          '*** MENSAJE PARA GESTION ***
          DataSAT = "b12256781234W00000011db2bch    " + _
                    "exec acl_aumento_real " + _
                    Mid$(DataC, 24, 3) + "," + Mid$(DataC, 1, 9) + "," + Mid$(DataC, 11, 11) + "," + Format$(Date, "yyyymmdd") + "," + Mid$(DataC, 27, 13) + "##"

          '*** MENSAJE PARA LA AGENDA ***
          DataSAT = DataSAT + "b12256781234W00000011telema    " + _
                    "exec p_put_visado " + _
                    Mid$(DataC, 1, 9) + ",'" + Format$(Date, "mm/dd/yyyy") + " " + Format$(Time, "hh:mm") + "'," + "000000000" + "," + Oficina + "," + TipoVisado + "##"

          '**** CIERRE TX GESTION ****
          DataSAT = DataSAT + IdUsuario + "2147" + "W" + Mid$(DataC, 1, 26) + Format$(Date, "yyyymmdd") + "##"
       
          PunteroTxSat.AddNew
          PunteroTxSat("TXSAT") = LlaveSAT + DataSAT
          PunteroTxSat.Update
          PunteroTxSat.MoveFirst
       End If
    End If
        
    Vista = "SAT_"
    Do While Not PunteroTxSat.EOF
       If PunteroTxSat("INTENTOS") < 8 Then
          Llave = PunteroTxSat("TXSAT")
          If TransmiteInformacion(Vista, Llave, True) Then
             'borra linea
             PunteroTxSat.Delete
          Else
             PunteroTxSat.Edit
             PunteroTxSat("RETSRM") = Mid$(CadenaDatos, 1, 2)
             PunteroTxSat("STATUS") = Mid$(CadenaDatos, 3, 2)
             PunteroTxSat("RETMSG") = Mid$(CadenaDatos, 5, 2)
             PunteroTxSat("INTENTOS") = PunteroTxSat("INTENTOS") + 1
             PunteroTxSat.Update
          End If
       End If
       PunteroTxSat.MoveNext
    Loop
    PunteroTxSat.Close
    PunteroTabPrd.Close
    PunteroVsdPrd.Close
    BaseEnvioSRM.Close
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

Private Function TransmiteInformacion(ByVal Vista As String, ByVal Llave As String, ByVal SeBuscaCuenta As Boolean) As Boolean
Dim Usuario As String
Dim NodoDestino As String * 4
Dim AplicacionDestino As String * 4
  TransmiteInformacion = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> "" Then
    Usuario = ""
    CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Llave)
    If Mid$(CadenaDatos, 1, 6) = "000000" Then
      TransmiteInformacion = True
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
  ' FI2026 - PunteroVistas.Index = PunteroVistas.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
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
