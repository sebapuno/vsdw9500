Attribute VB_Name = "ENVIAR"
Declare Function srmw8 Lib "c:\bin\f\srmw8.dll" (ByVal shost As String, ByVal sapli As String, ByVal smensaje As String, slargo_mensaje As Integer, ByVal sstatus As String, ByVal sfuncion As String, ByVal sctx As String, ByVal scontrol As String) As Integer
Global sd_error  As Integer

Dim Usuario As String
Dim LlaveUp As String
Dim DataUp As String
Dim Nodo_Dst As String
Dim Apl_Dst As String

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
  Retorno = srmw8(Sd_Nodo, Sd_Aplicacion, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)
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
  ErrorComunicacion.Operacion.Caption = " " + Modo
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

Function send_data(sd_uvo As String, Sd_Llave As String, sd_data As String, SD_LEE As Integer, sd_notfound_err As Integer, Kilos As String) As String
  Dim Sd_Contexto As String * 2
  Dim Sd_Status As String * 2
  Dim Sd_Funcion As String * 2
  Dim Sd_Control As String
  Dim Sd_Datos   As String * 8000
  Dim Sd_Largo   As Integer
  Dim sd_cod_ret As Integer
  Dim SD_CHECK   As String * 2

  Sd_Control = Space$(10000)
  If Kilos = "8K" Then
    Sd_Funcion = "08"
  Else
    Sd_Funcion = "01"
  End If
  Sd_Contexto = "00"
  Sd_Status = "00"
  Sd_Largo = Len(sd_uvo + Sd_Llave)
  Sd_Datos = sd_uvo + Sd_Llave + Chr$(0)
    
  sd_cod_ret = srmw8(Nodo_Dst, Apl_Dst, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)

  SD_CHECK = Mid$(Sd_Datos, 1, 2)
  If sd_cod_ret = 0 And Sd_Status = "00" Then
     If SD_CHECK <> "00" Then
        If (SD_CHECK = "96" And sd_notfound = True) Or SD_CHECK <> "96" Then
            Call error_srm(sd_cod_ret, Sd_Status, SD_CHECK, Nodo_Dst, Sd_Llave)
            sd_error = -1
        Else
            sd_error = 0
            If Mid$(Sd_Datos, 1, 2) = "00" Then
               If Apl_Dst = "SPL2" Then
                  Operacion = "C"
               Else
                  Operacion = "W"
               End If
            Else
               If Apl_Dst = "SPL2" Then
                  Operacion = "A"
               Else
                  Operacion = "S"
               End If
            End If
        End If
     Else
        sd_error = 0
        If SD_CHECK = "00" Then
            If Apl_Dst = "SPL2" Then
                Operacion = "C"
            Else
                Operacion = "W"
            End If
        Else
            If Apl_Dst = "SPL2" Then
                Operacion = "A"
            Else
                Operacion = "S"
            End If
        End If
     End If
  Else
     Call error_srm(sd_cod_ret, Sd_Status, SD_CHECK, Nodo_Dst, Sd_Llave)
     sd_error = -1
  End If
  
  If Not SD_LEE And sd_error = 0 Then
    Sd_Datos = Mid$(sd_uvo, 1, 12) + Operacion + Sd_Llave + sd_data
    Sd_Datos = Space$(8000)
    
    sd_cod_ret = srmw8(Nodo_Dst, Apl_Dst, Sd_Datos, Sd_Largo, Sd_Status, Sd_Funcion, Sd_Contexto, Sd_Control)
     
     SD_CHECK = Mid$(Sd_Datos, 1, 2)
     If sd_cod_ret = 0 And Sd_Status = "00" And SD_CHECK = "00" Then
        sd_error = 0
     Else
        Call error_srm(sd_cod_ret, Sd_Status, SD_CHECK, Nodo_Dst, Sd_Llave)
        sd_error = -1
    End If
  End If
  If sd_error = 0 Then
     send_data = "0000" + Trim$(Sd_Datos)
  Else
     send_data = Format$(sd_cod_ret, "00") + Sd_Status + Trim$(Sd_Datos)
  End If
End Function

Sub error_srm(retfun As Integer, st As String, RetSRM As String, Nodo As String, ERR_LLAVE As String)
Dim descripcion As String
     
     Select Case retfun
         Case 1
               descripcion = "[Error FBSS SRPI]"
         Case 2
               descripcion = "[Proceso Abortado por otra interrupción]"
         Case 3
               descripcion = "[Falló Obtención de variable de ambiente LAN]"
         Case 4
               descripcion = "[Falló apertura de Log File]"
         Case 5
               descripcion = "[Falló inicialización de comunicaciones]"
         Case 6
               descripcion = "[Mensaje aceptado, había respuesta previa]"
         Case 7
               descripcion = "[Envío de ejecución. Esperando Respuesta]"
         Case 9
               descripcion = "[Función incorrecta en llamado Padre-Hijo]"
         Case 10
               descripcion = "[Falló obtención de LUNUMBER]"
         Case 11
               descripcion = "[Falló obtención de TIMEOUT]"
         Case 21
               descripcion = "[Falló obtención de CICS_NAME]"
         Case 30
               descripcion = "[Falló obtención de NODENAME]"
         Case 41
               descripcion = "[Error en apertura de canal de comunicación]"
         Case 42
               descripcion = "[Error en escritura de canal de comunicación]"
         Case 43
               descripcion = "[Error en lectura de canal de comunicación]"
         Case 44
               descripcion = "[Error en cierre de canal de comunicación]"
         Case 45
               descripcion = "[Error en largo de la data de respuesta]"
         Case 90, 91
               descripcion = "[Error en inicialización de comunicación]"
         Case 92
               descripcion = "[Fallaron 3 intentos]"
         Case 93
               descripcion = "[Error en data SRM recibida]"
         Case 97
               descripcion = "[Error en carga de proceso externo]"
         Case 98
               descripcion = "[Error en contexto de comunicaciones Padre Hijo]"
         Case 99
               descripcion = "[Error en versión de Padre Hijo]"
     End Select

     If retfun <> 0 Then
         descripcion = descripcion + Chr(13)
     End If

     Select Case st
          Case "01"
               descripcion = descripcion + "[Servidor de comunicaciones inactivo en nodo destino]"
          Case "02"
               descripcion = descripcion + "[LU no definida como receptora en nodo destino]"
          Case "03"
               descripcion = descripcion + "[No existe aplicación en nodo destino]"
          Case "04"
               descripcion = descripcion + "[LU no definida en nodo]"
          Case "05"
               descripcion = descripcion + "[No existe programa asociado en nodo destino]"
          Case "07"
               descripcion = descripcion + "[No hay LU disponible]"
          Case "08"
               descripcion = descripcion + "[No hay LU disponible en nodo RE-RUTEADOR]"
          Case "09"
               descripcion = descripcion + "[TIMEOUT]"
          Case "10"
               descripcion = descripcion + "[Nodo de destino no existe en tabla de SRM]"
          Case "11"
               descripcion = descripcion + "[Aplicación de origen/destino no existe en tabla SRM]"
          Case "12"
               descripcion = descripcion + "[Tabla de contexto llena (SRM VAX)]"
          Case "13"
               descripcion = descripcion + "[Sesión inválida]"
          Case "14"
               descripcion = descripcion + "[Código contexto inválido]"
          Case "15"
               descripcion = descripcion + "[Aplicación de destino no activa]"
          Case "16"
               descripcion = descripcion + "[No existe PID de proceso originador]"
          Case "17"
               descripcion = descripcion + "[FILE TRANSFER en progeso]"
          Case "18"
               descripcion = descripcion + "[FILE TRANSFER terminado]"
          Case "19"
               descripcion = descripcion + "[Código desconocido en campo tipo mensaje]"
          Case "20"
               descripcion = descripcion + "[Mensaje de origen no tiene formato SRM]"
          Case "21"
               descripcion = descripcion + "[Mensaje de respuesta no tiene formato SRM]"
          Case "22"
               descripcion = descripcion + "[Enlace no autorizado a SRM]"
          Case "23"
               descripcion = descripcion + "[SRM inactivo en nodo local]"
          Case "24"
               descripcion = descripcion + "[NODO ALIAS no está en tabla SRM]"
          Case "25"
               descripcion = descripcion + "[Error en largo en data transmitida]"
          Case "26"
               descripcion = descripcion + "[Programa cancelado por problema no determinado]"
          Case "27"
               descripcion = descripcion + "[No hay memoria suficiente]"
          Case "28"
               descripcion = descripcion + "[Se rompió enlace]"
          Case "29"
               descripcion = descripcion + "[Nodo origen desconocido]"
          Case "30"
               descripcion = descripcion + "[PID inválido]"
          Case "31"
               descripcion = descripcion + "[SRM bloqueado por el operador]"
          Case "32"
               descripcion = descripcion + "[Nodo cerrado]"
          Case "33"
               descripcion = descripcion + "[Nodo origen no existe en tabla SRM]"
          Case "34"
               descripcion = descripcion + "[No hay LU's enviadoras en nodo de destino]"
          Case "35"
               descripcion = descripcion + "[Información de seguridad inconsistente]"
          Case "36"
               descripcion = descripcion + "[No autorizado]"
          Case "37"
               descripcion = descripcion + "[Error en comando CICS]"
          Case "38"
               descripcion = descripcion + "[Falló carga o enganche de PGM]"
          Case "39"
               descripcion = descripcion + "[Fuera de horario]"
          Case "40"
               descripcion = descripcion + "[No corresponde]"
          Case "41"
               descripcion = descripcion + "[Condición de excepción en manejo de archivos]"
          Case "42"
               descripcion = descripcion + "[Error en HOOK-LINK abortado]"
          Case "61"
               descripcion = descripcion + "[Generación de PID-OK]"
          Case "62"
               descripcion = descripcion + "[Se validará autorización]"
          Case "63"
               descripcion = descripcion + "[No se valida autorización]"
          Case "64"
               descripcion = descripcion + "[Solicita validación de PID]"
          Case "65"
               descripcion = descripcion + "[Seguridad completa OK]"
               
     End Select

     If st <> "00" Then
         descripcion = descripcion + Chr(13)
     End If
     
     Select Case RetSRM
         Case "53"
              descripcion = descripcion + "[Vista no existe]"
         Case "54"
              descripcion = descripcion + "[Operación no existe]"
         Case "55"
              descripcion = descripcion + "[Error en formato (clave o datos de input)]"
         Case "56"
              descripcion = descripcion + "[Código de transacción no numérica]"
         Case "57"
              descripcion = descripcion + "[Código de transacción no existe o no inicializada]"
         Case "58"
              descripcion = descripcion + "[Máximo de transacciones alcanzado, no se puede comenzar una nueva]"
         Case "59"
              descripcion = descripcion + "[Error al leer archivo de transacciones]"
         Case "60"
              descripcion = descripcion + "[Error al grabar archivo de transacciones]"
         Case "61"
              descripcion = descripcion + "[Error de dato en el Host]"
         Case "65"
              descripcion = descripcion + "[Operación no autorizada para la vista especificada]"
         Case "70"
              descripcion = descripcion + "[Debe inicializar clave]"
         Case "71"
              descripcion = descripcion + "[Clave incorrecta]"
         Case "72"
              descripcion = descripcion + "[Clave expirada]"
         Case "73"
              descripcion = descripcion + "[Grupo no existe]"
         Case "74"
              descripcion = descripcion + "[Perfil no existe]"
         Case "75"
              descripcion = descripcion + "[Usuario inactivo]"
         Case "76"
              descripcion = descripcion + "[Superó intentos permitidos para login]"
         Case "80"
              descripcion = descripcion + "[Registro Bloqueado]"
         Case "81"
              descripcion = descripcion + "[Tabla de locks está llena]"
         Case "82"
              descripcion = descripcion + "[Error en formato de datos]"
         Case "83"
              descripcion = descripcion + "[Registro no está bloqueado]"
         Case "84"
              descripcion = descripcion + "[Registro bloqueado por servidor]"
         Case "88"
              descripcion = descripcion + "[ABBEND]"
         Case "89"
              descripcion = descripcion + "[Registro duplicado]"
         Case "90"
              descripcion = descripcion + "[Archivo asociado a la vista no existe]"
         Case "91"
              descripcion = descripcion + "[Registro con clave duplicada]"
         Case "92"
              descripcion = descripcion + "[ILLOGIC]"
         Case "93"
              descripcion = descripcion + "[Requerimientos inválidos (INVREQ)]"
         Case "94"
              descripcion = descripcion + "[IOERR o error al abriri el archivo]"
         Case "95"
              descripcion = descripcion + "[LENGERR]"
         Case "96"
              descripcion = descripcion + "[Dato solicitado no se encontró]"
         Case "97"
              descripcion = descripcion + "[Archivos cerrados]"
         Case "98"
              descripcion = descripcion + "[NOTSPACE]"
         Case "99"
              descripcion = descripcion + "[BROWSE END]"
     End Select
     descripcion = descripcion + "Llave : [" + ERR_LLAVE + "]" + Chr$(13)
     descripcion = descripcion + "Nodo  : [" + Nodo + "]"

     MsgBox descripcion + Chr(13) + "RET: " + Format$(retfun, "00") + "   STATUS: " + st + "   RET_SRM: " + RetSRM, 16, "Problemas de Comunicación"
     
     End
End Sub


Function Prepara_Msg()
   
   Vista = NumVista
   If Environ$("DESA") = "DESA" Then
      Nodo_Dst = "UDES"
   Else
      Nodo_Dst = "USBT"
   End If
   Apl_Dst = "SD20"
   Usuario = IdUsuario + Vista + Modo
   LlaveUp = Llave
   Datado = Space(1500)

End Function

Function Trae_Infor(Manera As Integer)
   
   Dim Ok
   
   Do
      Prepara_Msg
      Datado = send_data(Usuario, LlaveUp, DataUp, True, True, "8K")
      If Mid$(Datado, 1, 6) = "000000" Then
         Ok = True
         If Not Separa(Manera) Then
            Exit Do
         End If
      Else
         Ok = False
         Exit Do
      End If
   Loop
   
   Trae_Infor = Ok

End Function

Function Separa(Manera As Integer)
   Dim Ok As Integer
                 
   Ok = False
   Separa = Ok
End Function
