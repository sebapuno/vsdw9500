Attribute VB_Name = "Bajada"
Private Declare Function Srmw32 Lib "C:\Bin\F\srmw32.dll" (ByVal Shost As String, ByVal SApli As String, ByVal SMensaje As String, SLargo_Mensaje As Long, ByVal SStatus As String, ByVal SFuncion As String, ByVal SCtx As String, ByVal SControl As String) As Integer
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
                                      "describiendo la siguiente información."
                                      
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
Public Function BajadaPRN(Nodo As String, Aplicacion As String, Llave As String, Retorno As String) As Integer
Dim Ret As Integer
Dim Status As String
Dim funcion As String * 2
Dim Contexto As String * 2
Dim CControl As String * 10000
Dim Ret_Dato As String
Dim Nodo_Dst As String * 4
Dim Buff_Data As String * 10000
Dim Largo_Msg As Long  ' FI2026 Integer->Long (Srmw32 32-bit)
  Status = "00"
  funcion = "01"
  CControl = ""
  Contexto = "00"
  Nodo_Dst = Nodo
  Buff_Data = Llave
  Largo_Msg = Len(Trim$(Buff_Data))
  Ret = Srmw32(Nodo_Dst, Aplicacion, Buff_Data, Largo_Msg, Status, funcion, Contexto, CControl)
  Ret_Dato = Mid(Buff_Data, 1, 2)
  Retorno = String$(6, 0)
  Retorno = Format$(Str$(Ret), "00") + Format$(Status, "00") + Ret_Dato
  If ((Ret = 0) And (Status = "00") And ((Ret_Dato = "00") Or (Ret_Dato = "96"))) Then
    BajadaPRN = True
  Else
    BajadaPRN = False
  End If
  Llave = Buff_Data
End Function
Public Function Llamada(ByVal Sd_Nodo As String, ByVal Sd_Aplicacion As String, ByVal Sd_Usuario As String, ByVal Sd_Llave As String, ByVal Kilos As String) As String
Dim Retorno As Integer
Dim Sd_Largo As Long  ' FI2026 Integer->Long (Srmw32 32-bit)
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
  
  If UCase(Logea) = "SI" Then
    Call registra_las_llamadas(Sd_Nodo, Sd_Aplicacion, Sd_Datos)
  End If
  
  Retorno = Srmw32(Sd_Nodo, Sd_Aplicacion, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)
  Sd_Status = Format$(Val(Sd_Status), "00")
  Sd_RetSRM = Mid$(Sd_Datos, 1, 2)
  Llamada = Format$(Retorno, "00") + Sd_Status + Sd_Datos
End Function

Public Function Llamada_ConReintentar(ByVal Sd_Nodo As String, ByVal Sd_Aplicacion As String, ByVal Sd_Usuario As String, ByVal Sd_Llave As String, ByVal Kilos As String) As String
Dim Retorno As Integer
Dim Sd_Largo As Long  ' FI2026 Integer->Long (Srmw32 32-bit)
Dim Sd_Datos As String * 8000
Dim Sd_Status As String * 2
Dim Sd_Control As String * 10000
Dim Sd_Funcion As String * 2
Dim Sd_RetSRM As String * 2
Dim Sd_Contexto As String * 2
Dim Reintentar As Boolean

Dim Entrada       As String
Dim Flag_Continua As String
Dim Num_registro  As Integer


  If Kilos = "8K" Then
    Sd_Funcion = "08"
  Else
    Sd_Funcion = "01"
  End If
  
  Flag_Continua = "N"
  Contador = 1
  
  Sd_Largo = Len(Sd_Usuario + Sd_Llave)
  Sd_Datos = Sd_Usuario + Sd_Llave + Chr$(0)
  Entrada = Sd_Datos
  Sd_Contexto = "00"
  Sd_Status = "00"
  
  GReintento = False
  Reintentar = True
  
  Do While Reintentar
    
    If UCase(Logea) = "SI" Then
      Call registra_las_llamadas(Sd_Nodo, Sd_Aplicacion, Sd_Datos)
    End If
    
    Retorno = Srmw32(Sd_Nodo, Sd_Aplicacion, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)
    Sd_Status = Format$(Val(Sd_Status), "00")
    Sd_RetSRM = Mid$(Sd_Datos, 1, 2)
    
    If Vista_Original = "PLAE" Or Vista_Original = "PLAP" Then
       Flag_Continua = Mid(Sd_Datos, 3, 1)
       If Flag_Continua = "S" Then
         Num_registro = Val(Mid(Sd_Datos, 4, 7))
         CadenaDatos = Sd_Datos
         Contador = Contador + Num_registro
         Llamada_ConReintentar = Mid(Entrada, 1, 13) & Format(Str(Contador), "0000000") & Mid(Entrada, 21)
         Sd_Datos = Llamada_ConReintentar
         GReintento = True
         Reintentar = True
         
       Else
        Llamada_ConReintentar = Format$(Retorno, "00") + Sd_Status + Sd_Datos
        CadenaDatos = Sd_Datos
        Reintentar = False
        GReintento = False
       End If
    Else
      Llamada_ConReintentar = Format$(Retorno, "00") + Sd_Status + Sd_Datos
      CadenaDatos = Sd_Datos
      Reintentar = False
      GReintento = False
    End If
    
    If (Retorno = 0) And (Sd_Status = "00") Then
      If Sd_RetSRM <> "00" Then
        If Sd_RetSRM = "99" Then
           Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 100), _
                    Sd_Usuario + Sd_Llave, Transaccion, ModoTransaccion)
        Else
           Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 80), _
                    Sd_Usuario + Sd_Llave, Transaccion, ModoTransaccion)
        End If
      Else
        If Vista_Original = "PLAE" Or Vista_Original = "PLAP" Then
          If FormateaSYBASE(Usuario_Aux, False, False) Then
             
          End If
        End If
      End If
    Else
      Call ErrorSRM(Sd_Nodo, Sd_Aplicacion, Retorno, Sd_Status, Mid$(Sd_Datos, 1, 100), _
                    Sd_Usuario + Sd_Llave, Transaccion, ModoTransaccion)
    End If
    
    If GReintento Then
      Reintentar = True
    End If
    
  Loop
   
End Function
