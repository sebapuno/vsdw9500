Attribute VB_Name = "Bajada"
Public Declare Function Srmw32 Lib "C:\Bin\F\srmw32.DLL" (ByVal Shost As String, ByVal SApli As String, ByVal SMensaje As String, SLargo_Mensaje As Long, ByVal SStatus As String, ByVal SFuncion As String, ByVal SCtx As String, ByVal SControl As String) As Integer
Public Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Private Sub ErrorSRM(ByVal Nodo As String, ByVal Aplicacion As String, ByVal RetFuncion As Integer, ByVal RetStatus As String, ByVal RetSRM As String, ByVal Llave As String, ByVal TipoTransaccion As String, ByVal Modo As String)
Dim Mensaje As String
Dim NumError As String
  ErrorComunicacion.Nodo.Caption = UCase(Nodo)
  ErrorComunicacion.Servidor.Caption = Aplicacion
  ErrorComunicacion.Operacion.Caption = " " + "Consulta"  'Modo
  If RetFuncion <> 0 Then
    NumError = "ERR" + Format$(RetFuncion, "0")
    Mensaje = SacaValorINI("RetFuncion", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.RetFuncion.Caption = " (" + Format$(RetFuncion, "00") + ") " + Mensaje
  Else
    ErrorComunicacion.RetFuncion.Caption = " (00)"
  End If
  If RetStatus <> "00" Then
    NumError = "ERR" + RetStatus
    Mensaje = SacaValorINI("Status", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.Status.Caption = " (" + RetStatus + ")" + " " + Mensaje
  Else
    ErrorComunicacion.Status.Caption = " (00)"
  End If
  If Mid$(RetSRM, 1, 2) <> "00" And Mid$(RetSRM, 1, 2) <> "99" Then
    NumError = "ERR" + Mid$(RetSRM, 1, 2)
    Mensaje = SacaValorINI("ServDatos", NumError, DIRECTORIO_01 + ARCHIVO_INI02)
    ErrorComunicacion.ServidorDatos.Caption = Mid(RetSRM, 1, 100)
  ElseIf Mid$(RetSRM, 1, 2) = "99" Then
    ErrorComunicacion.ServidorDatos.Caption = " (" + Mid$(RetSRM, 1, 2) + ") " + Mid$(RetSRM, 3, 90)
  Else
    ErrorComunicacion.ServidorDatos.Caption = " (00)"
  End If
  ErrorComunicacion.Llave.Caption = " " + Mid(Llave, 1, 50)
  ErrorComunicacion.Mensaje.Caption = "Se ha producido un Problema de Comunicaciones." + Chr(10) + _
                                      "Por favor Reporte esto a Central de Fallas (3030) " + _
                                      "describiendo la siguiente informaciÃƒÂ³n."
                                      
  ErrorComunicacion.Transaccion.Caption = " (" + Mid$(RetSRM, 1, 2) + ") " + Mensaje
                                      
  'Select Case TipoTransaccion
  '  Case "C"
  '    ErrorComunicacion.Transaccion.Caption = " Consulta"
  '  Case "M"
  '    ErrorComunicacion.Transaccion.Caption = " Commit"
   ' Case "R"
  '    ErrorComunicacion.Transaccion.Caption = " RollBack"
  '  Case "S"
  '    ErrorComunicacion.Transaccion.Caption = " Envio SAT"
  'End Select
  ErrorComunicacion.Show 1
End Sub

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
  Llamada = Format$(Retorno, "00") + Sd_Status + Sd_Datos
End Function

Public Function Llamada_ConReintentar(ByVal Sd_Nodo As String, ByVal Sd_Aplicacion As String, ByVal Sd_Usuario As String, ByVal Sd_Llave As String) As String
Dim Retorno As Integer
Dim Sd_Largo As Integer
Dim Sd_Datos As String * 8000
Dim Sd_Status As String * 2
Dim Sd_Control As String * 10000
Dim Sd_Funcion As String * 2
Dim Sd_RetSRM As String * 2
Dim Sd_Contexto As String * 2
Dim Reintentar As Boolean

  Sd_Funcion = "08"
  Sd_Largo = Len(Sd_Usuario + Sd_Llave)
  Sd_Datos = Sd_Usuario + Sd_Llave + Chr$(0)
  Sd_Contexto = "00"
  Sd_Status = "00"
  
  GReintento = False
  Reintentar = True
  
  Do While Reintentar
    Retorno = SRMW8(Sd_Nodo, Sd_Aplicacion, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)
    Sd_Status = Format$(Val(Sd_Status), "00")
    Sd_RetSRM = Mid$(Sd_Datos, 1, 2)
    Llamada_ConReintentar = Format$(Retorno, "00") + Sd_Status + Sd_Datos
    
    Reintentar = False
    If (Retorno = 0) And (Sd_Status = "00") Then
      If Sd_RetSRM <> "00" Then
        If Sd_RetSRM = "99" Then
           Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 100), _
                    Sd_Usuario + Sd_Llave, Transaccion, ModoTransaccion)
        Else
           Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 80), _
                    Sd_Usuario + Sd_Llave, Transaccion, ModoTransaccion)
        End If
      End If
    Else
      Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 100), _
                             Sd_Usuario + Sd_Llave, Transaccion, ModoTransaccion)
    End If
    
    If GReintento Then
      Reintentar = True
    Else
      Exit Do
    End If
  Loop
   
End Function
