Attribute VB_Name = "PROCVLD"
  Dim BaseVisado As Database
  
  Dim Ap_VsdPrd  As Recordset
  Dim Ap_TabPrd  As Recordset
  
  Global NumVista  As String * 4
  Global IdUsuario As String * 8
  Global GetAis    As String * 70
  Global Llave     As String
  Global Ap_modo   As Integer
  Global Datado    As String * 520
  Global Modo      As String

Function AbonoBatch()
   Call BloqueaReg
   Ap_VsdPrd.Edit
   Ap_VsdPrd("PRD_DATA") = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 59) + "S" + Mid$(Ap_VsdPrd("PRD_DATA"), 61)
   Ap_VsdPrd.Update
   End
End Function


Public Function Mod_11(V As String)

   Dim digito, suma As Integer
   Dim caracter As String
   Dim numerico As Boolean
   suma = 0
   For digito = Len(V) To 1 Step -1
       caracter = Mid$(V, digito, 1)
       If InStr("0123456789", caracter) > 0 Then
          numerico = True
       Else
          numerico = False
       End If
       If numerico Then
          suma = suma + Val(Mid$(V, digito, 1)) * ((Len(V) - digito) Mod 6 + 2)
       Else
       
          digito = 1
       End If
   Next
   Select Case suma Mod 11
      Case 0
           dv = "0"
      Case 1
           dv = "K"
      Case Else
           dv = LTrim(Str(11 - (suma Mod 11)))
   End Select
   Mod_11 = dv
End Function

Function Conv(DataPos As String, TipoDato As String) As Integer
   
   Dim PosCad, LargoCad As Integer
   
   PosCad = 0
   LargoCad = 0
   I = InStr(DataPos, ",")
   If Not I = 0 Then
      PosCad = Val(Mid$(DataPos, 1, I - 1))
      LargoCad = Val(Mid$(DataPos, I + 1, Len(DataPos) - I))
   End If
   If PosCad = 0 And LargoCad = 0 Then
      Conv = -1
   Else
      If TipoDato = "P" Then Conv = PosCad Else Conv = LargoCad
   End If

End Function
Public Function GrabaParametro(grupo As String, Llave As String, parametro As String, archivo As String) As Boolean
    Dim f As Integer, lpReturnedString As String
    
    lpReturnedString = String$(255, 0)
    parametro = parametro + Mid(lpReturnedString, Len(parametro) + 2)
    f = WritePrivateProfileString(ByVal grupo, ByVal Llave, ByVal parametro, ByVal archivo)
    GrabaParametro = True

End Function

Public Function ObtieneParametro(grupo As String, Llave As String, archivo As String) As String
    Dim f As Integer, lpReturnedString As String
    lpReturnedString = String$(255, 0)
    nSize% = Len(lpReturnedString)
    f = GetPrivateProfileString(grupo, ByVal Llave, "", lpReturnedString, nSize%, archivo)
    parametro = Left$(lpReturnedString$, f)

    ObtieneParametro = parametro

End Function

Function Ap_Visado()

   Dim Actividad As String * 1
   Dim Nom_Arch As String
   
   Dim FechaOtor As String
   Dim VarSiebel As String
   Dim Ret As Integer
   Dim p As Integer
   Dim nTO As String
   Dim DestTO As String
   
   ' FI2026 SHG - AisGetUsr (WAIS.DLL 16-bit) deshabilitado
   
   'Ret = AisGetUsr(GetAis)
   IdUsuario = Format(Mid(GetAis, 1, 8), "00000000")
   
   Set BaseVisado = OpenDatabase("C:\DATA\VISADO\VISADO.MDB")
   Set Ap_VsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
   ' FI2026 - Ap_VsdPrd.Index = Ap_VsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
   Set Ap_TabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
   Ap_TabPrd.Index = "TAB_PRD"
   
    ' OBTIENE LOS PARAMETROS
    Ap_Rut = ObtieneParametro("PARAMETROS", "PARAM1", "C:\DATA\VISADO\VISADO.INI")
    Ap_Producto = ObtieneParametro("PARAMETROS", "PARAM2", "C:\DATA\VISADO\VISADO.INI")
    Ap_Correlativo = ObtieneParametro("PARAMETROS", "PARAM3", "C:\DATA\VISADO\VISADO.INI")
    Ap_modo = ObtieneParametro("PARAMETROS", "PARAM4", "C:\DATA\VISADO\VISADO.INI")
    VarSiebel = ObtieneParametro("SIEBEL", "TIPOVSD", "C:\DATA\VISADO\VISADO.INI")
    
      Ap_TabPrd.Seek "=", Ap_Producto
      If Not Ap_TabPrd.NoMatch Then
         Ap_VsdPrd.Seek "=", Ap_Rut, Ap_Producto, Ap_Correlativo
         If Not Ap_VsdPrd.NoMatch Then
         
            nTO = Mid$(Ap_VsdPrd!PRD_DATA, 27, 3)
            DestTO = TraeDestino_TO(nTO)
            p = GrabaParametro("SIEBEL", "CRETO", DestTO, "C:\DATA\VISADO\VISADO.INI")
   
            'If VarSiebel <> "NORMAL" Then
            '  OperCRS = Mid$(Ap_VsdPrd!PRD_DATA, 40, 1)
            '  FechaOtor = Mid$(Ap_VsdPrd!PRD_DATA, 60 + 225, 8)
            '  FechaOtor = Format$(FechaOtor, "&&/&&/&&&&")
            '
            '  If Not ExisteUF(FechaOtor) Then
            '    Continuar = False
            '    MsgBox "No se encuentra el valor de la UF para la fecha de otorgamiento del crédito.", 16, "Error de Validación"
            '  Else
            '    Continuar = True
            '  End If
            'Else
            '  Continuar = True
            'End If
         
            If VerificaTX() Then
               End
            End If
         Else
            Ap_VsdPrd.Close
            Ap_TabPrd.Close
            BaseVisado.Close
            End
         End If
      Else
         Ap_VsdPrd.Close
         Ap_TabPrd.Close
         BaseVisado.Close
         End
      End If
End Function


Public Function TraeDestino_TO(numTO As String) As String

Dim LlaveTO As String
Dim DataTO As String
Dim NodoDestino As String
Dim Aplicacion As String
Dim Retorno As String

TraeDestino_TO = "NORMAL"

If IsNumeric(numTO) Then

  NodoDestino = NodoSybase("plt_emp01")
  Aplicacion = Mid$(NodoDestino, 5, 4)
  NodoDestino = Mid$(NodoDestino, 1, 4)

  LlaveTO = "b12256781234W0000001Rplt_emp01 exec sp_cre1217_trae_destino_to " & numTO
  
  Retorno = Llamada(NodoDestino, Aplicacion, LlaveTO, "8K", "C")

  If Mid$(Retorno, 1, 6) = "000000" Then
    If InStr(Retorno, "~") > 0 Then
      DataTO = Mid$(Retorno, 19)
      TraeDestino_TO = Mid$(DataTO, 1, InStr(DataTO, "~") - 1)
    End If
  End If

End If


End Function
Function NodoSybase(C_SY As String) As String
  Dim PunteroSy  As Recordset
  
  Set BaseEnvioSRM = OpenDatabase("C:\DATA\VISADO\ENVIOSRM.MDB")
  
  Set PunteroSy = BaseEnvioSRM.OpenRecordset("SYBASE", dbOpenTable)
  ' FI2026 - PunteroSy.Index = PunteroSy.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroSy.Seek "=", C_SY
  If Not PunteroSy.NoMatch Then
     If Environ("DESA") = "DESA" Then
        NodoSybase = PunteroSy("Sy_NodoDesa")
     Else
        NodoSybase = PunteroSy("Sy_NodoProd")
     End If
     NodoSybase = NodoSybase + PunteroSy("Sy_Aplicacion")
  Else
    NodoSybase = "XXXXXXXX"
  End If
  PunteroSy.Close
  
  BaseEnvioSRM.Close
End Function
Function BloqueaReg()
   NumVista = Ap_TabPrd("PRD_VISTA_DST")
   If Ap_modo = "1" Then
      Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 490) + "V"
   Else
      Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 26)
      Modo = "L"
      If Trae_Infor(0) Then
         Ap_VsdPrd.Edit
         Ap_VsdPrd("PRD_DATA") = Mid$(Datado, 7, 491)
         Ap_VsdPrd.Update
         Llave = Mid$(Datado, 7, 490) + "N"
      Else
         Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 490) + "N"
      End If
   End If
   'Ap_VsdPrd.Close
   'Ap_TabPrd.Close
   'BaseVisado.Close
   Modo = "W"
   If Trae_Infor(0) Then
      p = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")
      BloqueaReg = True
   Else
      BloqueaReg = False
   End If

End Function

Function Posee_TX(CtaUso As String)
    If Not Val(Mid$(CtaUso, 1, 10)) = 0 And Mid$(CtaUso, 11, 1) = "N" Then
       Posee_TX = True
    Else
       Posee_TX = False
    End If
End Function

Function VerificaTX()

   Dim DatoTx As String * 11
   
   If Ap_modo = 1 Then
      DatoTx = Mid$(Ap_VsdPrd("PRD_DATA"), 50, 11)
      If Posee_TX(DatoTx) Then
         If Caja_Abierta() Then
            FRMVLD03.Show
         Else
            FRMVLD02.Show
         End If

         VerificaTX = False
      Else
         Call BloqueaReg
         VerificaTX = True
      End If
   Else
      Call BloqueaReg
      VerificaTX = True
   End If

End Function

Function Caja_Abierta()

   Dim DATOS As String
   
   If Dir("C:\Data\Cajas\Totales.dat") = "TOTALES.DAT" Then
      Open "C:\Data\Cajas\Totales.dat" For Input As #1
      DATOS = Input$(31, #1)
      Close #1
      If Mid$(DATOS, 23, 1) = "A" And Mid$(DATOS, 1, 8) = IdUsuario Then
         Caja_Abierta = True
      Else
         Caja_Abierta = False
      End If
   Else
      Caja_Abierta = False
   End If
End Function
