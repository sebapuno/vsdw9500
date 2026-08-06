Attribute VB_Name = "Bajada"
Private Declare Function Srmw32 Lib "C:\Bin\F\srmw32.DLL" (ByVal Shost As String, ByVal SApli As String, ByVal SMensaje As String, SLargo_Mensaje As Long, ByVal SStatus As String, ByVal SFuncion As String, ByVal SCtx As String, ByVal SControl As String) As Integer
Public Function Llamada(ByVal Sd_Nodo As String, ByVal Sd_Aplicacion As String, ByVal Sd_Usuario As String, ByVal Sd_Llave As String, ByVal Kilos As String) As String
Dim Retorno As Integer
Dim Sd_Largo As Integer
Dim Sd_Datos As String * 8000
Dim Sd_Status As String * 2
Dim Sd_Control As String * 10000
Dim Sd_Funcion As String * 2
Dim Sd_RetSRM As String * 2
Dim Sd_Contexto As String * 2
  If Kilos = "8K" Then
    Sd_Funcion = "08"
  Else
    Sd_Funcion = "01"
  End If
  Sd_Largo = Len(Sd_Usuario + Sd_Llave)
  Sd_Datos = Sd_Usuario + Sd_Llave + Chr$(0)
  Sd_Contexto = "00"
  Sd_Status = "00"
  
  Retorno = SRMW8(Sd_Nodo, Sd_Aplicacion, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)
  Sd_Status = Format$(Val(Sd_Status), "00")
  Sd_RetSRM = Mid$(Sd_Datos, 1, 2)
  If (Retorno = 0) And (Sd_Status = "00") Then
    If InStr("00, 96", Sd_RetSRM) = 0 Then
      Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Sd_RetSRM, Sd_Usuario + Sd_Llave)
    End If
  Else
    Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Sd_RetSRM, Sd_Usuario + Sd_Llave)
  End If
  Llamada = Format$(Retorno, "00") + Sd_Status + Sd_Datos
End Function
Private Sub ErrorSRM(ByVal Nodo As String, ByVal Aplicacion As String, ByVal RetFuncion As Integer, ByVal RetStatus As String, ByVal RetSRM As String, ByVal Llave As String)
Dim Mensaje As String
Dim NumError As String
  ErrorComunicacion.Nodo.Caption = Nodo
  ErrorComunicacion.Servidor.Caption = Aplicacion
  ErrorComunicacion.Operacion.Caption = " " + Modo
  If RetFuncion <> 0 Then
    NumError = "ERR" + Format$(RetFuncion, "0")
    Mensaje = SacaValorINI("RetFuncion", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.RetFuncion.Caption = " (" + Format$(RetFuncion, "00") + ") " + Mensaje
  Else
    ErrorComunicacion.RetFuncion.Caption = " (" + Format$(RetFuncion, "00") + ")"
  End If
  If RetStatus <> "00" Then
    NumError = "ERR" + RetStatus
    Mensaje = SacaValorINI("Status", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.Status.Caption = " (" + RetStatus + ")" + " " + Mensaje
  Else
    ErrorComunicacion.Status.Caption = " (" + RetStatus + ")"
  End If
  If RetSRM <> "00" Then
    NumError = "ERR" + RetSRM
    Mensaje = SacaValorINI("ServDatos", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.ServidorDatos.Caption = " (" + RetSRM + ") " + Mensaje
  Else
    ErrorComunicacion.ServidorDatos.Caption = " (" + RetSRM + ")"
  End If
  ErrorComunicacion.Llave.Caption = " " + Llave
  ErrorComunicacion.Mensaje.Caption = "Se ha producido un Problema de Comunicaciones en la Validación de la " + _
                                                                  Mid$(PRODUCTOACCESADO, 7) + " " + _
                                                                  Mid$(PRODUCTOACCESADO, 1, 5) + ".  " + _
                                                                  "Por favor Reporte esto a Central de Fallas (3030) describiendo la " + _
                                                                  "siguiente información."
  ErrorComunicacion.nombreAplicacion.Caption = " " + Mid$(EJECUTABLE, 1, 5) + Producto + Mid$(EJECUTABLE, 9)
  ErrorComunicacion.Show 1
End Sub
