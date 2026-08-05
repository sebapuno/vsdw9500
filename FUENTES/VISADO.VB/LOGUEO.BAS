Attribute VB_Name = "Logueo"
Declare Function LInsert_2 Lib "c:\bin\f\waislog.dll" (ByVal Nodo As String, ByVal prall As String, ByVal sdall As String, ByVal trall As String, ByVal ctall As String, ByVal FLAGS As String, ByVal autorizo As String, ByVal Resultado As String, ByVal perError As String, ByVal sdoError As String, ByVal logdata As String, ByVal Largo As Integer, ByVal hora_S As String, ByVal ind As String, ByVal funcion As Integer, ByVal msg As String) As Integer

Const MODOLOGUEO = 2

Sub LoguearVisado(Opcion As String, SOp As Integer, RutOper As String, RutCliente As String, Result As String, RetSRM As String, RetStat As String, Data As String, ProdLog As String)
Dim Err2 As String * 4
Dim Err1 As String * 4
Dim FLAGS As String * 16
Dim Indice As String * 8
Dim Buffer As String * 2000
Dim Llave1 As String * 32
Dim Llave2 As String * 32
Dim Llave3 As String * 32
Dim Llave4 As String * 32
Dim SubOpc As String * 8
Dim NodoLog As String * 4
Dim PunteroTabPrd As Recordset
  
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  PunteroTabPrd.Seek "=", ProdLog
  
  HoraLogueo$ = Format$(Now, "hhmmss")
  Canal$ = Environ("Sistema")
  Canal$ = Mid$(Canal$ + String$(8 - Len(Canal$), " "), 1, 8)
  IdProg$ = Mid$(String$(8 - Len(Ejecutable), " ") + Ejecutable, 1, 8)
  SubOpc = Mid$(Format$(SOp, "00000000"), 1, 8)
  NodoLog = UCase$(Environ("Srm_NO"))
  RutCliente = IIf(RutCliente = "", "123456789", RutCliente)
  FLAGS = "NNSS" + String(12, " ")
  Llave1 = Mid$(Canal$ + IdProg$ + Opcion + SubOpc, 1, 32)
  Llave2 = "VSD" + String$(23, " ") + HoraLogueo$
  Llave3 = Mid$(Environ("Oficina"), 1, 3) + String$(9, " ") + _
           Right(String$(10, "0") + Trim$(RutCliente), 10) + String$(10, "0")
  Llave4 = ProdLog + " " + Mid$(PunteroTabPrd("Prd_Descrip"), 1, 24) + " " + Mid$(Environ("Oficina"), 1, 3)
  Err1 = Mid$(RetSRM, 1, 2) + String$(2, " ")
  Err2 = Mid$(RetStat, 1, 2) + String$(2, " ")
  Data = Data + String$(256 - Len(Data), " ") + HoraLogueo$
  Indice = String(8, " ")
  Buffer = String(2000, " ")
  Respuesta% = LInsert_2(NodoLog, Llave1, Llave2, Llave3, Llave4, FLAGS, RutOper, Result, _
                                       Err1, Err2, Data, 262, HoraInicioSesion, _
                                       Indice, MODOLOGUEO, Buffer)
  If (Respuesta% > 0) And (Respuesta% <> 32) Then
    Call ErrorDeLogueo("V", Respuesta%)
  End If
  PunteroTabPrd.Close

End Sub
Sub LoguearTransaccion(RutOper As String, Datos As String)
Dim Err2 As String * 4
Dim Err1 As String * 4
Dim Hora As String
Dim Data As String
Dim FLAGS As String * 16
Dim Indice As String
Dim Buffer As String * 2000
Dim Llave1 As String * 32
Dim Llave2 As String * 32
Dim Llave3 As String * 32
Dim Llave4 As String * 32
Dim NodoLog As String * 4
Dim Operacion As String * 1
  Err1 = Mid$(Datos, 292, 4)
  Err2 = Mid$(Datos, 296, 4)
  Data = Mid$(Datos, 6, 262)
  Hora = Mid$(Datos, 448, 16)
  FLAGS = Mid$(Datos, 432, 16)
  Indice = "VSDPROD1"
  Llave1 = Mid$(Datos, 304, 32)
  Llave2 = Mid$(Datos, 336, 32)
  Llave3 = Mid$(Datos, 368, 32)
  Llave4 = Mid$(Datos, 400, 32)
  NodoLog = Environ("Srm_NO")
  Operacion = Mid$(Datos, 462, 1)
  Respuesta% = LInsert_2(NodoLog, Llave1, Llave2, Llave3, Llave4, FLAGS, RutOper, Operacion, _
                                       Err1, Err2, Data, 262, Hora, Indice, MODOLOGUEO, Buffer)
  If (Respuesta% > 0) And (Respuesta% <> 32) Then
    Call ErrorDeLogueo("T", Respuesta%)
  End If
End Sub
Private Sub ErrorDeLogueo(ByVal CualTitulo As String, ByVal Retorno As Integer)
Dim Mensaje As String
Dim NumError As String
  Select Case CualTitulo
    Case "T"
      ErrorLogueo.Caption = "Error en Logueo de la Transacción"
    Case "V"
      ErrorLogueo.Caption = "Error en Logueo del Visado"
  End Select
  Mensaje = SacaValorINI("TipoLogueo", MODOLOGUEO, DIRECTORIO_01 + ARCHIVO_INI02)
  ErrorLogueo.Tipo.Caption = " (" + LTrim$(Str(MODOLOGUEO)) + ") " + Mensaje
  NumError = "ERR" + Format$(Retorno, "00")
  Mensaje = SacaValorINI("RetLogueo", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
  ErrorLogueo.Glosa.Caption = " (" + Format$(Retorno, "00") + ") " + Mensaje
  ErrorLogueo.Show 1
End Sub


Function registra_las_llamadas(xnodo As String, xaplica As String, xdata As String)

Dim archivo_LECTURA As String
Dim Archivo_DESTINO As String
Dim XGRABAR As String
Dim MiRegistrox As String * 1000 ' Registro

   
Archivo_DESTINO = "d:\VSD_LOG.TXT"
       
Open Archivo_DESTINO For Append As #2
MiRegistrox = ""
MiRegistrox = Date & " " & Time & " " & xnodo & " " & xaplica & " " & xdata
Print #2, MiRegistrox
Close #2
       
End Function
