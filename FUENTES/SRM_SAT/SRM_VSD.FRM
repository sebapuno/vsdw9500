VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1185
   ClientLeft      =   2100
   ClientTop       =   2745
   ClientWidth     =   3195
   Height          =   1590
   Left            =   2040
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   3195
   Top             =   2400
   Visible         =   0   'False
   Width           =   3315
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Llave As String
Dim Datado As String * 1500
Dim IdUsuario As String * 8
Dim NumVista As String * 4
Dim BaseEnvioSRM As Database
Dim BaseVisado   As Database
Dim ProdPrinc As String * 3
Dim MensajeBack As String
Function Es_Sybase(C_SY As String) As String
Dim PunteroSy  As Recordset

  Set PunteroSy = BaseEnvioSRM.OpenRecordset("SYBASE", dbOpenTable)
  ' FI2026 - PunteroSy.Index = PunteroSy.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroSy.Seek "=", C_SY
  If Not PunteroSy.NoMatch Then
     If Ambiente = "DESA" Then
        Es_Sybase = PunteroSy("Sy_NodoDesa")
     Else
        Es_Sybase = PunteroSy("Sy_NodoProd")
     End If
     Es_Sybase = Es_Sybase + PunteroSy("Sy_Aplicacion")
  Else
    Es_Sybase = "XXXXXXXX"
  End If
  PunteroSy.Close
End Function

Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  Else
    PoneValorINI = False
  End If
End Function
Private Function ObtieneNumeroTransaccion(ByVal NodoDestino As String, ByVal Aplicacion As String) As String
Dim NroTx As String * 6
Dim LlaveUp As String
  NroTx = BuscaNumeroTransaccion(NumVista)
  If NroTx = SINENCONTRAR Then
    Usuario = IdUsuario + NumVista + "I"
    LlaveUp = ""
    Datado = Llamada(NodoDestino, Aplicacion, LlaveUp, "8K", "C")
    If Mid$(Datado, 1, 6) = "000000" Then
      NroTx = Mid$(Datado, 7, 6)
      Call GrabaNumeroTransaccion(NumVista, NroTx)
    End If
  End If
  ObtieneNumeroTransaccion = NroTx
End Function
Private Sub GrabaNumeroTransaccion(ByVal Tx_Vista As String, ByVal Tx_NumeroTx As String)
Dim PunteroTx As Recordset
  Set PunteroTx = BaseEnvioSRM.OpenRecordset("Tx", dbOpenTable)
  PunteroTx.AddNew
  PunteroTx("Tx_Vista") = Tx_Vista
  PunteroTx("Tx_NroTX") = Tx_NumeroTx
  PunteroTx.Update
  PunteroTx.Close
End Sub
Private Function DeterminaNodo(ByVal NumeroVista As String, Aplicacion As String, Cad_Sybase As String) As String
Dim PunteroVista As Recordset
Dim PunteroPrd   As Recordset
Dim Resp_Sy      As String
Dim PunteroAlias   As Recordset
Dim Busca_Ofi    As String
Dim Oficina As String
Dim IniDesa As String
Dim NodoDesa As String

IniDesa = SacaValorINI("VISADO", "DESA", DIRECTORIO_01 + ARCHIVO_INI01)
If IniDesa = "DESA" Or Environ$("DESA") = "DESA" Then
 Ambiente = "DESA"
 NodoDesa = SacaValorINI("VISADO", "DESANODO", DIRECTORIO_01 + ARCHIVO_INI01)
End If

  Resp_Sy = Es_Sybase(Cad_Sybase)
  If Resp_Sy <> "XXXXXXXX" Then
     DeterminaNodo = Mid$(Resp_Sy, 1, 4)
     Aplicacion = Mid$(Resp_Sy, 5, 4)
  Else
     Set PunteroVista = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
     ' FI2026 - PunteroVista.Index = PunteroVista.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
     PunteroVista.Seek "=", NumeroVista
     If PunteroVista.NoMatch Then
       DeterminaNodo = SINNODO
     Else
       
       If Ambiente = "DESA" Then
         If PunteroVista("Vta_NodoDesa") = "U???" Then
           DeterminaNodo = NodoDesa 'LCase$(Environ$("SRM_NS"))
         Else
           DeterminaNodo = PunteroVista("Vta_NodoDesa")
         End If
       Else
         Busca_Ofi = SacaValorINI("PRODUCTO", "RETORNO", DIRECTORIO_01 + ARCHIVO_INI01)

         If Busca_Ofi = "OK" And PunteroVista("Vta_BuscaOfi") = "S" And NumeroVista <> "1249" And NumeroVista <> "4249" Then
           Oficina = SacaValorINI("PRODUCTO", "OFICINA", DIRECTORIO_01 + ARCHIVO_INI01)
           DeterminaNodo = BuscaNodoOfi(Oficina)
           
         ElseIf PunteroVista("Vta_NodoProd") = "U???" Then
           DeterminaNodo = LCase$(Environ$("SRM_NS"))
           
         Else
           DeterminaNodo = PunteroVista("Vta_NodoProd")
         End If
       End If
       
       ' Analiza si el producto ya es 2000
       Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
       
       ProductoRelacionado = ProdPrinc
       Set PunteroAlias = BaseVisado.OpenRecordset("TAB_SYBALIAS", dbOpenTable)
       ' FI2026 - PunteroAlias.Index = PunteroAlias.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
       PunteroAlias.Seek "=", ProdPrinc
       If Not PunteroAlias.NoMatch Then
         ProductoRelacionado = PunteroAlias!syb_prod_nvo
       End If
       If PunteroAlias.NoMatch Then
         Set PunteroAlias = BaseVisado.OpenRecordset("TAB_ALIAS", dbOpenTable)
         ' FI2026 - PunteroAlias.Index = PunteroAlias.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
         PunteroAlias.Seek "=", ProdPrinc
         If Not PunteroAlias.NoMatch Then
           ProductoRelacionado = PunteroAlias!alias_prod_nuev
         End If
       End If
       PunteroAlias.Close
       
       Aplicacion = PunteroVista("Vta_Aplicacion")
       
       BaseVisado.Close
     End If
     PunteroVista.Close
  End If
End Function
Private Function BuscaNumeroTransaccion(ByVal Vista As String) As String
Dim PunteroTx As Recordset
  BuscaNumeroTransaccion = SINENCONTRAR
  Set PunteroTx = BaseEnvioSRM.OpenRecordset("Tx", dbOpenTable)
  ' FI2026 - PunteroTx.Index = PunteroTx.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTx.Seek "=", Vista
  If Not PunteroTx.NoMatch Then
    BuscaNumeroTransaccion = PunteroTx("Tx_NroTx")
  End If
  PunteroTx.Close
End Function
Private Sub Commit()
Dim IdTx As String * 6
Dim Nodo As String
Dim PunteroTx As Recordset
Dim Aplicacion As String
  Set PunteroTx = BaseEnvioSRM.OpenRecordset("Tx", dbOpenTable)
  If PunteroTx.RecordCount > 0 Then
    PunteroTx.MoveLast
    Do While Not PunteroTx.BOF
      NumVista = PunteroTx("TX_Vista")
      Nodo = DeterminaNodo(NumVista, Aplicacion, "")
      Llave = IdUsuario + NumVista + "M"
      IdTx = PunteroTx("TX_NroTx")
      If TransmiteDatos(Nodo, IdTx, "M", Aplicacion) Then
        PunteroTx.Delete
        PunteroTx.MovePrevious
      Else
        BaseEnvioSRM.Execute ("INSERT INTO Tx_Hst SELECT Tx.* FROM Tx;")
        Exit Do
      End If
    Loop
  End If
  PunteroTx.Close
End Sub
Private Function TransmiteDatos(ByVal Nodo As String, ByVal IdTx As String, ByVal TipoTransaccion As String, ByVal Aplicacion As String) As Boolean
Dim LlaveUp As String
  TransmiteDatos = False
  Usuario = Mid$(Llave, 1, 12)
  LlaveUp = Mid$(Llave, 14)
  If IdTx <> SINTRANSACCION Then
    Select Case Mid$(Llave, 13, 1)
      Case "E"
        Usuario = Usuario + "D" + IdTx
      Case "M"
        Usuario = Usuario + "M" + IdTx
      Case "R"
        Usuario = Usuario + "R" + IdTx
      Case "S"
        Usuario = Usuario + "A" + IdTx
      Case "W"
        Usuario = Usuario + "K" + IdTx 'PARA BLOQUEAR REGISTRO
        Datado = Llamada(Nodo, Aplicacion, LlaveUp, "8K", TipoTransaccion)
        Usuario = Mid$(Llave, 1, 12)
        Usuario = Usuario + "C" + IdTx
    End Select
  Else
'hacer arreglo
    Usuario = Mid$(Llave, 1, 13)
  End If
  
  Datado = Llamada(Nodo, Aplicacion, LlaveUp, "8K", TipoTransaccion)
  If Mid$(Datado, 1, 6) = "000000" Then
    TransmiteDatos = True
  End If
End Function
Private Sub RollBack()
Dim IdTx As String * 6
Dim Nodo As String
Dim Aplicacion As String
Dim PunteroTx As Recordset
  Set PunteroTx = BaseEnvioSRM.OpenRecordset("Tx", dbOpenTable)
  If PunteroTx.RecordCount > 0 Then
    MsgBox "Se RealizarÃ¡ RollBack a la TransacciÃ³n", 0 + 64, "Anular TransacciÃ³n"
    PunteroTx.MoveLast
    Do While Not PunteroTx.BOF
      NumVista = PunteroTx("Tx_Vista")
      IdTx = PunteroTx("Tx_NroTx")
      Nodo = DeterminaNodo(NumVista, Aplicacion, "")
      Llave = IdUsuario + NumVista + "R"
      If TransmiteDatos(Nodo, IdTx, "R", Aplicacion) Then
        PunteroTx.Delete
        PunteroTx.MovePrevious
      Else
        BaseEnvioSRM.Execute ("INSERT INTO Tx_Hst SELECT Tx.* FROM Tx;")
        Exit Do
      End If
    Loop
  End If
  PunteroTx.Close
End Sub
Private Sub Form_Load()
Dim IdTx As String * 6
Dim NodoDestino As String
Dim Rut         As String
Dim PosIni As Integer
Dim GetAis As String * 70
Dim PosFin As Integer
Dim Mensaje As String
Dim Aplicacion As String
Dim VisadoOK As Integer
Dim PunteroMensaje As Recordset
Dim D_TRANSAC As String * 4
Dim MsgBack_Temp As String


  ' FI2026 SHG - AisGetUsr (WAIS.DLL 16-bit) deshabilitado


  'Ret = AisGetUsr(GetAis)
  IdUsuario = Format$(Mid(GetAis, 1, 8), "00000000")
  
  
  'Para Enviar a SAT se enviaran las tx con el rut del cliente y no con el
  'usuario logeado
  Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Rut = Mid(Rut, 2, 8)
    
  PosIni = 1
  PosFin = 1
  VisadoOK = True
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set PunteroMensaje = BaseEnvioSRM.OpenRecordset("Mensaje")
  
  ProdPrinc = PunteroMensaje("Cod_ProdPrinc")
  
  If PunteroMensaje("SAT") Then
     ' VIA SAT
     NodoDestino = DeterminaNodo("SAT_", Aplicacion, "")
     D_TRANSAC = PunteroMensaje("TX_SAT")
     IdTx = "XXXXXX"
     'Modif 08/2005 JSA a Pedido de Juan Ferrari
     'Mensaje = IdUsuario + D_TRANSAC + "G" + "0000" + "VSD_"
     Mensaje = Rut + D_TRANSAC + "G" + "0000" + "VSD_"
     Select Case D_TRANSAC
        Case "0001"       '* VISADO TARJETA
            Mensaje = Mensaje + "Visado Tarjeta " + "99" + PunteroMensaje("Data")
        Case "1001"
            Mensaje = Mensaje + "Elimina Tarjeta" + "10" + PunteroMensaje("Data")
        Case "0002"       '* VISADO CREDITO PERSONAS
            Mensaje = Mensaje + "Visado CRP     " + "10" + PunteroMensaje("Data")
        Case "0003"       '* VISADO CREDITO EMPRESAS
            Mensaje = Mensaje + "Visado CRE     " + "10" + PunteroMensaje("Data")
        'Modif JSA 07/2005
        'Para trasmitir CRF o CRB se usara 0006
        'Case "0004"       '* VISADO CREDITO PERSONAS FUSION
            'Mensaje = Mensaje + "Visado CRB     " + "10" + PunteroMensaje("Data")
        'Case "1004"       '* RECHAZO
        '    Mensaje = Mensaje + "Rechazo CRB    " + "10" + PunteroMensaje("Data")
        'Case "0005"       '* VISADO CREDITO PERSONAS CON DESEMBOLSO EN LINEA
            'Mensaje = Mensaje + "Visado CRF     " + "10" + PunteroMensaje("Data")
        Case "0006"       '* VISADO CREDITO PERSONAS CON DESEMBOLSO EN LINEA
            Mensaje = Mensaje + "Visado CRF-CRB " + "10" + PunteroMensaje("Data")
  
     End Select
     Llave = Mensaje
     If Not TransmiteDatos(NodoDestino, IdTx, "S", Aplicacion) Then
        VisadoOK = False
     End If
     
     If VisadoOK And Not IsNull(PunteroMensaje!Directo) Then
       If InStr(PunteroMensaje!Directo, "##") > 0 Then
         If Not TransmiteDirecto(PunteroMensaje!Directo) Then
           'VisadoOK = False
         End If
       End If
     End If
     
  Else
     BaseEnvioSRM.Execute "DELETE * FROM Tx;"
     Do While PosFin > 0
       PosFin = InStr(PosIni, PunteroMensaje!Data, "##")
       If PosFin > 0 Then
         Mensaje = Mid$(PunteroMensaje!Data, PosIni, PosFin - PosIni)
         If Mensaje <> "NULL" Then
           If Mid$(Mensaje, 11, 8) = "PLT_PYME" Then
              NumVista = Mid$(Mensaje, 21, 4)
           Else
              If Mid$(Mensaje, 1, 5) = "BCH  " Then ' Factoring
                 NumVista = Mid$(Mensaje, 42, 4)
              Else
                 NumVista = Mid$(Mensaje, 9, 4)
              End If
           End If
           If Mid$(Mensaje, 1, 4) = "ADT0" Then NumVista = "ADT0"
           NodoDestino = DeterminaNodo(NumVista, Aplicacion, Mid$(Mensaje, 22, 10))
           
           ' Prepara data simple para CTD1 - Cta Remunerada
           If Mid$(Mensaje, 9, 4) = "ADT0" Then
              Mensaje = Trim$(Mid$(Mensaje, 14)) '* sin rut+vista+operacion
           End If
           
           ' Codigos Magicos Para Banca Privada Traspaso Cuenta Corriente
           ' Clasica - Privada - Clasica
           If NumVista = "Y5R5" Or NumVista = "X5R5" Then
              NumVista = "4391"
              NodoDestino = Right(Mensaje, 4)
              Mensaje = Mid$(Mensaje, 1, 8) & "4391" & Mid$(Mensaje, 13, Len(Mensaje) - 4)
           End If
           
           ' Codigos Magicos Para Banca Privada Traspaso Cuenta Corriente
           ' Clasica - Privada - Clasica
           ' HOMOLOGACION
           If NumVista = "Y6R5" Or NumVista = "X6R5" Then
              NumVista = "6391"
              NodoDestino = Right(Mensaje, 4)
              Mensaje = Mid$(Mensaje, 1, 8) & "6391" & Mid$(Mensaje, 13, Len(Mensaje) - 4)
           End If
           
           ' Codigos Magicos Para Banca Privada Regional
           If Mid$(NumVista, 1, 3) = "ZRZ" Then
              Select Case Mid$(NumVista, 4, 1)
                Case "1"
                  NumVista = "4391"
                Case "2"
                  NumVista = "4400"
                Case "3"
                  NumVista = "1358"
                Case "4"
                  NumVista = "5391"
                Case "5"
                  NumVista = "6391"  '* HOMOLOGACION
              End Select
              NodoDestino = Right(Mensaje, 4)
              Mensaje = Mid$(Mensaje, 1, 8) & NumVista & Mid$(Mensaje, 13, Len(Mensaje) - 4)
           End If
           
           ' Prepara data simple para CTD1 o CTR1 - Cta Con Intereses
           If Mid$(NumVista, 1, 4) = "CTD1" Or Mid$(NumVista, 1, 4) = "CTR1" Then
              Mensaje = Trim$(Mid$(Mensaje, 14)) '* sin rut+vista+operacion
           End If
           ' Prepara data simple para CRP0 - Credito en Cuotas FusiÃ³n
           If Mid$(NumVista, 1, 4) = "CRP0" Then
              Mensaje = Trim$(Mid$(Mensaje, 14)) '* sin rut+vista+operacion
           End If
           
           MsgBack_Temp = ""
           If NumVista = "LCD2" And Mid$(Mensaje, 14, 2) = "03" Then
              MsgBack_Temp = Mid$(Mensaje, 1, 13) + "04" + Mid$(Mensaje, 16) + "##"
              Mensaje = Mid$(Mensaje, 14)
           End If
           
           If Mid$(NumVista, 1, 4) = "SCS2" Then
              Mensaje = Mid$(Mensaje, 1, 8) + Mid$(Mensaje, 13) '+ "##"
           End If
           
           If Mid$(NumVista, 1, 4) = "SEM1" Then
              Mensaje = Mid$(Mensaje, 1, 8) + Mid$(Mensaje, 14) '+ "##"
           End If
                   
           ' Fin Procedimiento
           If NodoDestino <> SINNODO Then
              If Aplicacion = "SD20" Then
                IdTx = ObtieneNumeroTransaccion(NodoDestino, Aplicacion)
                If IdTx <> SINENCONTRAR Then
                   Llave = Mensaje
                   If Not TransmiteDatos(NodoDestino, IdTx, "C", Aplicacion) Then
                      Call RollBack
                      VisadoOK = False
                      Exit Do
                   End If
                Else
                   Call RollBack
                   VisadoOK = False
                   Exit Do
                End If
              Else
                IdTx = SINTRANSACCION
                Llave = Mensaje
                If Not TransmiteDatos(NodoDestino, IdTx, "C", Aplicacion) Then
                  If Trim$(MsgBack_Temp) <> "" Then
                    MensajeBack = MensajeBack + MsgBack_Temp
                  End If
                  Call RollBack
                  Call RollBackOtros '* Solo rollback distinto al SD20
                  VisadoOK = False
                  Exit Do
                End If
              End If
           Else
             MsgBox "No Existe la Vista " + NumVista + " en la Tabla " + Chr(34) + "Vistas" + Chr(34), _
                           0 + 16, "TransmisiÃ³n Cancelada"
           End If
         End If
       End If
       PosIni = PosFin + 2
     Loop
  End If
  If VisadoOK Then
    If Not PunteroMensaje("SAT") Then Call Commit
    Retorno% = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
  End If
  PunteroMensaje.Close
  BaseEnvioSRM.Close
  End
End Sub


Public Function Interpreta(Cadena As String, DatoBuscado As String) As Integer
Dim PosicionComa As Integer

  PosicionComa = InStr(Cadena, ",")
  If PosicionComa <> 0 Then
    If DatoBuscado = SIMBOLOPOSICION Then
      Interpreta = Val(Mid(Cadena, 1, PosicionComa - 1))
    Else
      Interpreta = Val(Mid(Cadena, PosicionComa + 1, Len(Cadena)))
    End If
  Else
    Interpreta = -1
  End If
End Function





Public Sub RollBackOtros()

Dim PIni As Integer
Dim PFin As Integer
Dim Mens As String
Dim Aplic As String
Dim Retorno_Ok As Integer

PIni = 1
PFin = IIf(Trim$(MensajeBack) <> "", InStr(PIni, MensajeBack, "##"), 0)

Do While PFin > 0
  PFin = InStr(PIni, MensajeBack, "##")
  If PFin > 0 Then
    Mens = Mid$(MensajeBack, PIni, PFin - PIni)
    NodoDestino = DeterminaNodo(Mid$(Mens, PIni + 8, 4), Aplic, "")
    
    If Mid$(Mens, 9, 4) = "LCD2" Then
      Llave = Mid$(Mens, 14)
    Else
      Llave = Mens
    End If
    
    Retorno_Ok = 0
    While Retorno_Ok < 3
      If Not TransmiteDatos(NodoDestino, SINTRANSACCION, "C", Aplic) Then
        Retorno_Ok = Retorno_Ok + 1
      End If
    Wend
  End If
  PIni = PFin + 2
Loop

MensajeBack = ""

End Sub



Public Function TransmiteDirecto(DataMensaje As String) As Boolean
Dim IdTx As String * 6
Dim NodoDestino As String
Dim PosIni As Integer
Dim GetAis As String * 70
Dim PosFin As Integer
Dim Mensaje As String
Dim Aplicacion As String
Dim VisadoOK As Integer
Dim PunteroMensaje As Recordset
Dim D_TRANSAC As String * 4
Dim MsgBack_Temp As String

BaseEnvioSRM.Execute "DELETE * FROM Tx;"

TransmiteDirecto = True
PosIni = 1
PosFin = 1

Do While PosFin > 0
  PosFin = InStr(PosIni, DataMensaje, "##")
  If PosFin > 0 Then
    Mensaje = Mid$(DataMensaje, PosIni, PosFin - PosIni)
    If Mensaje <> "NULL" Then
      
      NodoDestino = DeterminaNodo(NumVista, Aplicacion, Mid$(Mensaje, 22, 10))
                 
      ' Prepara data simple para CTD1 o CTR1 - Cta Con Intereses
      If Mid$(NumVista, 1, 4) = "CTD1" Or Mid$(NumVista, 1, 4) = "CTR1" Then
        Mensaje = Trim$(Mid$(Mensaje, 14)) '* sin rut+vista+operacion
      End If
      ' Prepara data simple para CRP0 - Credito en Cuotas FusiÃ³n
      If Mid$(NumVista, 1, 4) = "CRP0" Then
        Mensaje = Trim$(Mid$(Mensaje, 14)) '* sin rut+vista+operacion
      End If
           
      If Mid$(NumVista, 1, 4) = "SCS2" Then
        Mensaje = Mid$(Mensaje, 1, 8) + Mid$(Mensaje, 13) + "##"
      End If
           
      ' Fin Procedimiento
      If NodoDestino <> SINNODO Then
        If Aplicacion = "SD20" Then
          IdTx = ObtieneNumeroTransaccion(NodoDestino, Aplicacion)
          If IdTx <> SINENCONTRAR Then
            Llave = Mensaje
            If Not TransmiteDatos(NodoDestino, IdTx, "C", Aplicacion) Then
              Call RollBack
              TransmiteDirecto = False
              Exit Do
            End If
          Else
            Call RollBack
            TransmiteDirecto = False
            Exit Do
          End If
        Else
          IdTx = SINTRANSACCION
          Llave = Mensaje
          If Not TransmiteDatos(NodoDestino, IdTx, "C", Aplicacion) Then
            If Trim$(MsgBack_Temp) <> "" Then
              MensajeBack = MensajeBack + MsgBack_Temp
            End If
            Call RollBack
            TransmiteDirecto = False
            Exit Do
          End If
        End If
      Else
        MsgBox "No Existe la Vista " + NumVista + " en la Tabla " + Chr(34) + "Vistas" + Chr(34), _
                   0 + 16, "TransmisiÃ³n Cancelada"
      End If
    End If
  End If
  PosIni = PosFin + 2
Loop

End Function
