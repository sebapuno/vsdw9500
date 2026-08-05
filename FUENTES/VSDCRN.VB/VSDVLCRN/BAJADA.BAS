Attribute VB_Name = "Bajada"
Private Declare Function SRMW8 Lib "C:\Bin\F\SRMW8.DLL" (ByVal Shost As String, ByVal SApli As String, ByVal SMensaje As String, SLargo_Mensaje As Integer, ByVal SStatus As String, ByVal SFuncion As String, ByVal SCtx As String, ByVal SControl As String) As Integer



Public Function Llamada(ByVal Sd_Nodo As String, ByVal Sd_Aplicacion As String, ByVal Sd_Llave As String, ByVal Kilos As String, ByVal Transaccion As String) As String
Dim Retorno As Integer
Dim Sd_Largo As Integer
Dim Sd_Datos As String * 8000
Dim Sd_Status As String * 2
Dim Sd_Control As String * 10000
Dim Sd_Funcion As String * 2
Dim Sd_RetSRM As String * 2
Dim Sd_Contexto As String * 2
Dim ModoTransaccion As String * 1
  If (Kilos = "8K") And (InStr("0000, 00D1", Sd_Nodo) = 0) Then
    Sd_Funcion = "08"
  Else
    Sd_Funcion = "01"
  End If
  Sd_Largo = Len(Usuario + Sd_Llave)
  Sd_Datos = Usuario + Sd_Llave + Chr$(0)
  Sd_Contexto = "00"
  Sd_Status = "00"
  ModoTransaccion = Mid$(Sd_Datos, 13, 1)
  Clipboard.SetText Sd_Nodo + Sd_Aplicacion + Sd_Datos
  Retorno = SRMW8(Sd_Nodo, Sd_Aplicacion, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)
  Sd_Status = Format$(Val(Sd_Status), "00")
  Sd_RetSRM = Mid$(Sd_Datos, 1, 2)
  If (Retorno = 0) And (Sd_Status = "00") Then
    If Sd_RetSRM <> "00" Then
      If Sd_RetSRM = "99" Then
         Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 260), _
                    Usuario + Sd_Llave, Transaccion, ModoTransaccion)
      Else
         Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 260), _
                    Usuario + Sd_Llave, Transaccion, ModoTransaccion)
      End If
    End If
  Else
    Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 150), _
                             Usuario + Sd_Llave, Transaccion, ModoTransaccion)
  End If
  Llamada = Format$(Retorno, "00") + Sd_Status + Trim$(Sd_Datos)
End Function
Private Sub ErrorSRM(ByVal Nodo As String, ByVal Aplicacion As String, ByVal RetFuncion As Integer, ByVal RetStatus As String, ByVal RetSRM As String, ByVal Llave As String, ByVal TipoTransaccion As String, ByVal Modo As String)
Dim Mensaje As String
Dim NumError As String
  ErrorComunicacion.Nodo.Caption = Nodo
  ErrorComunicacion.Servidor.Caption = Aplicacion
  ErrorComunicacion.OPERACION.Caption = " " + Modo
  If RetFuncion <> 0 Then
    NumError = "ERR" + Format$(RetFuncion, "0")
    Mensaje = ObtieneParametro("RetFuncion", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.RetFuncion.Caption = " (" + Format$(RetFuncion, "00") + ") " + Mensaje
  Else
    ErrorComunicacion.RetFuncion.Caption = " (00)"
  End If
  If RetStatus <> "00" Then
    NumError = "ERR" + RetStatus
    Mensaje = ObtieneParametro("Status", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.Status.Caption = " (" + RetStatus + ")" + " " + Mensaje
  Else
    ErrorComunicacion.Status.Caption = " (00)"
  End If
  
  If Aplicacion = "CRP0" Then
    NumError = Mid$(RetSRM, 1, 2)
    If NumError = "50" And Val(Mid$(RetSRM, 48, 5)) > 0 And Val(Mid$(RetSRM, 113, 2)) > 0 Then
      Mensaje = Mid$(RetSRM, 115, 60)
      ErrorComunicacion.ServidorDatos.Caption = " (" + NumError + ") " + Mensaje
    Else
      ErrorComunicacion.ServidorDatos.Caption = " (" + NumError + ") " + Space(4) + Mid(RetSRM, 1, 40)
    End If

  ElseIf Mid$(RetSRM, 1, 2) <> "00" And Mid$(RetSRM, 1, 2) <> "99" And Aplicacion <> "LCD2" Then
    NumError = "ERR" + Mid$(RetSRM, 1, 2)
    Mensaje = ObtieneParametro("ServDatos", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.ServidorDatos.Caption = " (" + Mid$(RetSRM, 1, 2) + ") " + Mensaje + Space(4) + Mid(RetSRM, 1, 40)
    
  ElseIf Mid$(RetSRM, 1, 2) = "99" Or Aplicacion = "LCD2" Then
    ErrorComunicacion.ServidorDatos.Caption = " (" + Mid$(RetSRM, 1, 2) + ") " + Mid$(RetSRM, 3, 90)
  
  Else
    ErrorComunicacion.ServidorDatos.Caption = " (00)"
  End If
  ErrorComunicacion.Llave.Caption = " " + Mid(Llave, 1, 50)
  ErrorComunicacion.Mensaje.Caption = "Se ha producido un Problema de Comunicaciones en la Transmisión " + _
                                                                  "de un Producto.  Por favor Reporte esto a Central de Fallas (3030) " + _
                                                                  "describiendo la siguiente información."
  Select Case TipoTransaccion
    Case "C"
      ErrorComunicacion.Transaccion.Caption = " Consulta"
    Case "M"
      ErrorComunicacion.Transaccion.Caption = " Commit"
    Case "R"
      ErrorComunicacion.Transaccion.Caption = " RollBack"
    Case "S"
      ErrorComunicacion.Transaccion.Caption = " Envio SAT"
  End Select
  ErrorComunicacion.Show 1
End Sub
