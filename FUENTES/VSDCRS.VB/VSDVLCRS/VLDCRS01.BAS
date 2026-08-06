Attribute VB_Name = "PROCVLD"
Option Explicit

Global Ap_VsdPrd  As Recordset
   
Global NumVista  As String * 4
Global IdUsuario As String * 8
Global GetAis    As String * 70
Global Llave     As String
Global Ap_modo   As Integer
Global Modo      As String

Dim Datado As String * 1500
Dim BaseEnvioSRM As Database
Dim BaseVisado   As Database

Dim Usuario As String
Dim Sd_Llave As String
Dim DataUp As String
Dim Nodo_Dst As String
Dim Apl_Dst As String
Global DataPatron As String

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const ARCHIVO_INI01 = "Visado.INI"
Function AbonoBatch()
   
   If Ap_modo = "1" Then
      Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 490 + 59) + "V"
   End If
   
   Ap_VsdPrd.Edit
   Ap_VsdPrd("PRD_DATA") = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 59) + "S" + Mid$(Ap_VsdPrd("PRD_DATA"), 61)
   Ap_VsdPrd.Update
   
   Ap_VsdPrd.Close
   BaseVisado.Close
   
   End
   
End Function


Function BloqueaReg()
   
   If Ap_modo = "1" Then
      Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 490 + 59) + "V"
   End If
   BloqueaReg = True
  
End Function
Public Sub BorraRegistro(bData As String)

Dim largo As Integer
Dim RecVsdPrd  As Recordset
  
largo = Len(bData)

Set RecVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE " + _
                           "Mid$(Prd_Data, 1, " & CStr(largo) & ")='" & bData & "'", dbOpenDynaset)
  
If RecVsdPrd.RecordCount > 0 Then
  RecVsdPrd.Delete
End If
  
RecVsdPrd.Close

End Sub

Public Function Mod_11(V As String)

   Dim digito, suma As Integer
   Dim caracter As String
   Dim numerico As Boolean
   Dim dv As String
   
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
   Dim i As Integer
   
   PosCad = 0
   LargoCad = 0
   i = InStr(DataPos, ",")
   If Not i = 0 Then
      PosCad = Val(Mid$(DataPos, 1, i - 1))
      LargoCad = Val(Mid$(DataPos, i + 1, Len(DataPos) - i))
   End If
   If PosCad = 0 And LargoCad = 0 Then
      Conv = -1
   Else
      If TipoDato = "P" Then Conv = PosCad Else Conv = LargoCad
   End If

End Function
Public Function GrabaParametro(Grupo As String, Llave As String, parametro As String, archivo As String) As Boolean
    Dim f As Integer, lpReturnedString As String
    
    lpReturnedString = String$(255, 0)
    parametro = parametro + Mid(lpReturnedString, Len(parametro) + 2)
    f = WritePrivateProfileString(ByVal Grupo, ByVal Llave, ByVal parametro, ByVal archivo)
    GrabaParametro = True

End Function

Public Function ObtieneParametro(Grupo As String, Llave As String, archivo As String) As String
    Dim f As Integer, lpReturnedString As String
    Dim nSize%
    Dim parametro
    
    lpReturnedString = String$(255, 0)
    nSize% = Len(lpReturnedString)
    f = GetPrivateProfileString(Grupo, ByVal Llave, "", lpReturnedString, nSize%, archivo)
    parametro = Left$(lpReturnedString$, f)

    ObtieneParametro = parametro

End Function


Function Ap_Visado()

Dim Actividad As String * 1
Dim Nom_Arch As String
Dim Ap_Rut  As String
Dim Ap_Producto As String
Dim Ap_Correlativo As String
Dim Continuar As Boolean
Dim MtoIni As String
Dim CapMO As String
Dim MtoCred As String
Dim TipCam As String
Dim OperCRS As String
Dim FchVigSeg As String
Dim Ret
Dim Encabezado As String
Dim LlaveProducto As String
Dim cantidad
Dim i#, j#, k#
Dim p
Dim FechaOtor As String
Dim VarSiebel As String
Dim nTO As String
Dim DestTO As String

' FI2026 SHG - AisGetUsr (WAIS.DLL 16-bit) deshabilitado

'Ret = AisGetUsr(GetAis)
IdUsuario = Format(Mid(GetAis, 1, 8), "00000000")
 
' OBTIENE LOS PARAMETROS
Ap_Rut = ObtieneParametro("PARAMETROS", "PARAM1", "C:\DATA\VISADO\VISADO.INI")
Ap_Producto = ObtieneParametro("PARAMETROS", "PARAM2", "C:\DATA\VISADO\VISADO.INI")
Ap_Correlativo = ObtieneParametro("PARAMETROS", "PARAM3", "C:\DATA\VISADO\VISADO.INI")
Ap_modo = ObtieneParametro("PARAMETROS", "PARAM4", "C:\DATA\VISADO\VISADO.INI")
VarSiebel = ObtieneParametro("SIEBEL", "TIPOVSD", "C:\DATA\VISADO\VISADO.INI")
  
Set BaseVisado = OpenDatabase("C:\DATA\VISADO\VISADO.MDB")
Set Ap_VsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
' FI2026 - Ap_VsdPrd.Index = Ap_VsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
Ap_VsdPrd.Seek "=", Ap_Rut, Ap_Producto, Ap_Correlativo
   
If Not Ap_VsdPrd.NoMatch Then
       
 Encabezado = Mid$(Ap_VsdPrd!PRD_DATA, 1, 6)
 LlaveProducto = Mid$(Ap_VsdPrd!PRD_DATA, 10, 30)
 OperCRS = Mid$(Ap_VsdPrd!PRD_DATA, 40, 1)
 FechaOtor = Mid$(Ap_VsdPrd!PRD_DATA, 60 + 225, 8)
 FechaOtor = Format$(FechaOtor, "&&/&&/&&&&")
 
 nTO = Mid$(Ap_VsdPrd!PRD_DATA, 60 + 27, 3)
 DestTO = TraeDestino_TO(nTO)
 p = GrabaParametro("SIEBEL", "CRETO", DestTO, "C:\DATA\VISADO\VISADO.INI")
 
 '* Calcula CrÃ©dito
 'MtoIni = Mid$(Ap_VsdPrd!PRD_DATA, 60 + 61, 17)
 'CapMO = Mid$(Ap_VsdPrd!PRD_DATA, 60 + 78, 17)
 'TipCam = Mid$(Ap_VsdPrd!PRD_DATA, 60 + 214, 11)
 'MtoCred = CalculaMontoCreditoPesos(OperCRS, MtoIni, CapMO, TipCam)
  
 'If OperCRS = "R" Then
  If ProductoAsociadoEncontrado(Encabezado + "SC1" + LlaveProducto, "1", "39") Then
    MtoCred = Mid$(Mid$(DataPatron, 69), 187 + 12, 13)
    FchVigSeg = Mid$(Mid$(DataPatron, 69), 442, 10)
    FchVigSeg = Mid$(FchVigSeg, 4, 3) + Mid$(FchVigSeg, 1, 3) + Mid$(FchVigSeg, 7, 4)
                        
    If AgregarMontoCredito(Encabezado + "SC1" + LlaveProducto, "1", "39", MtoCred) Then
       cantidad = Val(Mid$(DataPatron, 61 + 6, 2))
       Continuar = True
       '* SEGUROS
       For i# = 2 To cantidad
         If Not ProductoAsociadoEncontrado(Encabezado + "SC" + Trim$(Str(i#)) + LlaveProducto, "1", "39") Then
           MsgBox "Registros del seguro son incompletos.", 16, "Error de ValidaciÃ³n"
           Continuar = False
           Exit For
         ElseIf Not AgregarMontoCredito(Encabezado + "SC" + Trim$(Str(i#)) + LlaveProducto, "1", "39", MtoCred) Then
           Continuar = False
           Exit For
         End If
       Next i#
       '* CLIENTE PRODUCTO
       If Continuar Then
         For j# = 1 To cantidad
           If Not ProductoAsociadoEncontrado(Encabezado + "SR" + Trim$(Str(j#)) + LlaveProducto, "1", "39") Then
             MsgBox "Registros del seguro son incompletos.", 16, "Error de ValidaciÃ³n"
             Continuar = False
             Exit For
           End If
         Next j#
       End If
       '* DIRECCIONES
       If Continuar Then
         For k# = 1 To cantidad
           If Not ProductoAsociadoEncontrado(Encabezado + "SD" + Trim$(Str(k#)) + LlaveProducto, "1", "39") Then
             MsgBox "Registros del seguro son incompletos.", 16, "Error de ValidaciÃ³n"
             Continuar = False
             Exit For
           Else
             Call ModificaDatosDireccion(Encabezado + "SD" + Trim$(Str(k#)) + LlaveProducto, "1", "39")
           End If
         Next k#
       End If
       
       If Not IsNumeric(MtoCred) Or InStr(MtoCred, ".") <> 0 Then
         Continuar = False
         MsgBox "El Monto de Cobertura del Seguro es incorrecto.", 16, "Error de ValidaciÃ³n"
       End If
       
       If Not IsDate(FchVigSeg) Or DateDiff("d", FchVigSeg, Now) > 30 Then
         Continuar = False
         MsgBox "La fecha de vigencia del seguro expirÃ³.", 16, "Error de ValidaciÃ³n"
       End If
       
       If VarSiebel <> "NORMAL" And Continuar Then
         If Not ExisteUF(FechaOtor) Then
           Continuar = False
           MsgBox "No se encuentra el valor de la UF para la fecha de otorgamiento del crÃ©dito.", 16, "Error de ValidaciÃ³n"
         End If
       End If
       
       If Continuar Then
          p = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")
       End If
         
    End If
  Else
    MsgBox "No se encuentran registros del seguro asociado.", 16, "Error de ValidaciÃ³n"
  End If
  
 'Else '* IF OperCRS<>R  '** BORRAR REGISTROS DE SRn y SDn.
   
  ' For j# = 1 To 7
  '   If ProductoAsociadoEncontrado(Encabezado + "SR" + Trim$(Str(j#)) + LlaveProducto, "1", "39") Then
  '      BorraRegistro (Encabezado + "SR" + Trim$(Str(j#)) + LlaveProducto)
  '   End If
   
 '    If ProductoAsociadoEncontrado(Encabezado + "SD" + Trim$(Str(j#)) + LlaveProducto, "1", "39") Then
 '       BorraRegistro (Encabezado + "SD" + Trim$(Str(j#)) + LlaveProducto)
 '    End If
 '  Next j#
   
 '  p = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")
   
 'End If '* IF OperCRS="R"
 
End If
   
Ap_VsdPrd.Close
BaseVisado.Close
End
   
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

Private Sub ModificaDatosDireccion(ByVal PatronBusqueda As String, ByVal Posicion As String, ByVal largo As String)

Dim FiltroVsdPrd  As Recordset
Dim DataDir As String

DataPatron = ""
    
Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE " + _
                           "Mid$(Prd_Data, " + Posicion + ", " + CStr(Len(PatronBusqueda)) + ")='" + PatronBusqueda + "'", dbOpenDynaset)
  
If FiltroVsdPrd.RecordCount > 0 Then

  DataDir = FiltroVsdPrd("Prd_Data")
  '**Obliga a DirecciÃ³n Postal
  DataDir = Mid$(DataDir, 1, 60 + 10) + "3" + Mid$(DataDir, 72)
   
  FiltroVsdPrd.Edit
  FiltroVsdPrd("Prd_Data") = DataDir
  FiltroVsdPrd.Update
 
End If
  
FiltroVsdPrd.Close
  
End Sub

Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal Posicion As String, ByVal largo As String) As Boolean

Dim FiltroVsdPrd  As Recordset

DataPatron = ""
  
ProductoAsociadoEncontrado = False
  
Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE " + _
                           "Mid$(Prd_Data, " + Posicion + ", " + CStr(Len(PatronBusqueda)) + ")='" + PatronBusqueda + "';", dbOpenSnapshot)
  
If FiltroVsdPrd.RecordCount > 0 Then
  DataPatron = FiltroVsdPrd("Prd_Data")
  ProductoAsociadoEncontrado = True
End If
  
FiltroVsdPrd.Close
  
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
      DatoTx = Mid$(Ap_VsdPrd!PRD_DATA, 50 + 60, 11) ' 50 , 11) CRS
      If Posee_TX(DatoTx) Then
         If Caja_Abierta() Then
            FRMVLD03.Show
         Else
            FRMVLD02.Show
         End If
         VerificaTX = False
         
      Else
         If Ap_modo = "1" Then
           Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 490 + 59) + "V"
         End If
         VerificaTX = True
      End If
   Else
      If Ap_modo = "1" Then
         Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 490 + 59) + "V"
      End If
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

Public Function CalculaMontoCreditoPesos(OperCred As String, MtoInicial As String, MtoCapMO As String, MtoTipoCambio As String) As String

Dim MontoIni As String
Dim CapitalMO As String
Dim TipoCambio As String
Dim MontoCreMO As String
Dim MontoFlot

MontoIni = Trim$(Decodifica(MtoInicial, 0, 17))
CapitalMO = Trim$(Decodifica(MtoCapMO, 0, 17))
TipoCambio = Trim$(Decodifica(MtoTipoCambio, 0, 11))

If OperCred = "R" And Val(CapitalMO) <> 0 Then
   MontoCreMO = Format$(MontoIni - CapitalMO, String(17, "0"))
Else
   MontoCreMO = Format$(MontoIni, String(17, "0"))
End If

If Val(Mid$(TipoCambio, 1, 7)) <> 1 Then
  
  MontoCreMO = Mid$(MontoCreMO, 1, 13) + "," + Mid$(MontoCreMO, 14)
  TipoCambio = Mid$(TipoCambio, 1, 7) + "," + Mid$(TipoCambio, 8)

  MontoFlot = CDbl(MontoCreMO) * CDbl(TipoCambio)
  CalculaMontoCreditoPesos = Format$(CLng(MontoFlot), String(13, "0"))

Else
  CalculaMontoCreditoPesos = Mid$(MontoCreMO, 1, 13)
End If


End Function
Public Function Decodifica(Numero As String, Posicion As Integer, largo As Integer) As String
Dim Signo As String * 1
Dim Numerito As String
Dim LargoTotal  As Integer
Dim PosicionInicio As Integer
  If Numero <> "" Then
    If largo <= 0 Then
      PosicionInicio = 1
      LargoTotal = Posicion - 1
    Else
      PosicionInicio = Posicion + 1
      LargoTotal = largo - 1
    End If
    If InStr("{ABCDEFGHI", Mid$(Numero, Len(Numero), 1)) <> 0 Then
      Signo = ""
    Else
      Signo = "-"
    End If
    Select Case Mid$(Numero, Len(Numero), 1)
      Case "{", "}"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "0"
      Case "A", "J"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "1"
      Case "B", "K"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "2"
      Case "C", "L"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "3"
      Case "D", "M"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "4"
      Case "E", "N"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "5"
      Case "F", "O"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "6"
      Case "G", "P"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "7"
      Case "H", "Q"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "8"
      Case "I", "R"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "9"
    End Select
    If largo <= 0 Then
      Decodifica = Signo + Format$(Numerito, "0000000000000")
    Else
      Decodifica = Signo + Format$(Mid(Numero, 1, Posicion), "0000000000000") + Numerito
    End If
  Else
    Decodifica = " " + String(13, "0")
  End If
End Function


Public Function AgregarMontoCredito(ByVal PatronBusqueda As String, ByVal Posicion As String, ByVal largo As String, MontoCredito As String) As Boolean

'** Si se encuentra el registro le agrega al final el dato
'** monto del crÃ©dito en pesos.

Dim FiltroVsdPrd  As Recordset

AgregarMontoCredito = False
  
Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE " + _
                           "Mid$(Prd_Data, " + Posicion + ", " + largo + ")='" + PatronBusqueda + "'", dbOpenDynaset)
  
If FiltroVsdPrd.RecordCount > 0 Then
  FiltroVsdPrd.Edit
  FiltroVsdPrd("Prd_Data") = Mid$(FiltroVsdPrd("Prd_Data"), 1, 60 + 637) + MontoCredito
  FiltroVsdPrd.Update
  AgregarMontoCredito = True
End If
  
FiltroVsdPrd.Close
End Function

Public Function ExisteUF(FechaOtorgamiento As String) As Boolean

Dim LlaveUp As String
Dim DataUp As String
Dim NodoDestino As String
Dim Aplicacion As String
Dim Retorno As String

'  FechaOtorgamiento="05/10/2005"
ExisteUF = False

If IsDate(FechaOtorgamiento) Then

  NodoDestino = NodoSybase("plt_emp01")
  Aplicacion = Mid$(NodoDestino, 5, 4)
  NodoDestino = Mid$(NodoDestino, 1, 4)

  LlaveUp = "b12256781234W00000012plt_emp01 exec sp_busca_unidad_fomento '" & FechaOtorgamiento & "'"

  Retorno = Llamada(NodoDestino, Aplicacion, LlaveUp, "8K", "C")

  If Mid$(Retorno, 1, 6) = "000000" Then
    If Mid$(Retorno, 19, 2) = "SI" Then
       ExisteUF = True
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
