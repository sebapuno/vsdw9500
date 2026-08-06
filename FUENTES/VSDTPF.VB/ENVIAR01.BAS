Attribute VB_Name = "ENVIAR"
Declare Function Srmw32 Lib "c:\bin\f\srmw32.dll" (ByVal shost As String, ByVal sapli As String, ByVal smensaje As String, slargo_mensaje As Long, ByVal sstatus As String, ByVal sfuncion As String, ByVal sctx As String, ByVal scontrol As String) As Integer
Global sd_error  As Integer

Dim Usuario As String
Dim LlaveUp As String
Dim DataUp As String
Dim Nodo_Dst As String
Dim Apl_Dst As String
Function send_data(SD_UVO As String, SD_LLAVE As String, SD_DATA As String, SD_LEE As Integer, sd_notfound_err As Integer, kilos As String) As String
  Dim SD_CONTEXTO As String * 2
  Dim SD_STATUS As String * 2
  Dim SD_FUNCION As String * 2
  Dim SD_CONTROL As String
  Dim SD_DATOS   As String * 8000
  Dim SD_LARGO   As Integer
  Dim SD_COD_RET As Integer
  Dim SD_CHECK   As String * 2

  SD_CONTROL = Space$(10000)
  If kilos = "8K" Then
    SD_FUNCION = "08"
  Else
    SD_FUNCION = "01"
  End If
  SD_CONTEXTO = "00"
  SD_LARGO = Len(SD_UVO + SD_LLAVE)
  SD_DATOS = SD_UVO + SD_LLAVE + Chr$(0)
  SD_STATUS = "00"
  SD_COD_RET = SRMW8(Nodo_Dst, Apl_Dst, SD_DATOS, SD_LARGO, SD_STATUS, SD_FUNCION, SD_CONTEXTO, SD_CONTROL)

  SD_CHECK = Mid$(SD_DATOS, 1, 2)
  If SD_COD_RET = 0 And SD_STATUS = "00" Then
     If SD_CHECK <> "00" Then
        If (SD_CHECK = "96" And sd_notfound = True) Or SD_CHECK <> "96" Then
            Call error_srm(SD_COD_RET, SD_STATUS, SD_CHECK, Nodo_Dst, SD_LLAVE)
            sd_error = -1
        Else
            sd_error = 0
            If Mid$(SD_DATOS, 1, 2) = "00" Then
               If Apl_Dst = "SPL2" Then
                  OPERACION = "C"
               Else
                  OPERACION = "W"
               End If
            Else
               If Apl_Dst = "SPL2" Then
                  OPERACION = "A"
               Else
                  OPERACION = "S"
               End If
            End If
        End If
     Else
        sd_error = 0
        If SD_CHECK = "00" Then
            If Apl_Dst = "SPL2" Then
                OPERACION = "C"
            Else
                OPERACION = "W"
            End If
        Else
            If Apl_Dst = "SPL2" Then
                OPERACION = "A"
            Else
                OPERACION = "S"
            End If
        End If
     End If
  Else
     Call error_srm(SD_COD_RET, SD_STATUS, SD_CHECK, Nodo_Dst, SD_LLAVE)
     sd_error = -1
  End If
  
  If Not SD_LEE And sd_error = 0 Then
    SD_DATOS = Mid$(SD_UVO, 1, 12) + OPERACION + SD_LLAVE + SD_DATA
    SD_DATOS = Space$(8000)
    
    SD_COD_RET = SRMW8(Nodo_Dst, Apl_Dst, SD_DATOS, SD_LARGO, SD_STATUS, SD_FUNCION, SD_CONTEXTO, SD_CONTROL)
     
     SD_CHECK = Mid$(SD_DATOS, 1, 2)
     If SD_COD_RET = 0 And SD_STATUS = "00" And SD_CHECK = "00" Then
        sd_error = 0
     Else
        Call error_srm(SD_COD_RET, SD_STATUS, SD_CHECK, Nodo_Dst, SD_LLAVE)
        sd_error = -1
    End If
  End If
  If sd_error = 0 Then
     send_data = "0000" + Trim$(SD_DATOS)
  Else
     send_data = Format$(SD_COD_RET, "00") + SD_STATUS + Trim$(SD_DATOS)
  End If
End Function

Sub error_srm(retfun As Integer, st As String, retsrm As String, NODO As String, ERR_LLAVE As String)
Dim descripcion As String
     
     Select Case retfun
         Case 1
               descripcion = "[Error FBSS SRPI]"
         Case 2
               descripcion = "[Proceso Abortado por otra interrupciÃ³n]"
         Case 3
               descripcion = "[FallÃ³ ObtenciÃ³n de variable de ambiente LAN]"
         Case 4
               descripcion = "[FallÃ³ apertura de Log File]"
         Case 5
               descripcion = "[FallÃ³ inicializaciÃ³n de comunicaciones]"
         Case 6
               descripcion = "[Mensaje aceptado, habÃ­a respuesta previa]"
         Case 7
               descripcion = "[EnvÃ­o de ejecuciÃ³n. Esperando Respuesta]"
         Case 9
               descripcion = "[FunciÃ³n incorrecta en llamado Padre-Hijo]"
         Case 10
               descripcion = "[FallÃ³ obtenciÃ³n de LUNUMBER]"
         Case 11
               descripcion = "[FallÃ³ obtenciÃ³n de TIMEOUT]"
         Case 21
               descripcion = "[FallÃ³ obtenciÃ³n de CICS_NAME]"
         Case 30
               descripcion = "[FallÃ³ obtenciÃ³n de NODENAME]"
         Case 41
               descripcion = "[Error en apertura de canal de comunicaciÃ³n]"
         Case 42
               descripcion = "[Error en escritura de canal de comunicaciÃ³n]"
         Case 43
               descripcion = "[Error en lectura de canal de comunicaciÃ³n]"
         Case 44
               descripcion = "[Error en cierre de canal de comunicaciÃ³n]"
         Case 45
               descripcion = "[Error en largo de la data de respuesta]"
         Case 90, 91
               descripcion = "[Error en inicializaciÃ³n de comunicaciÃ³n]"
         Case 92
               descripcion = "[Fallaron 3 intentos]"
         Case 93
               descripcion = "[Error en data SRM recibida]"
         Case 97
               descripcion = "[Error en carga de proceso externo]"
         Case 98
               descripcion = "[Error en contexto de comunicaciones Padre Hijo]"
         Case 99
               descripcion = "[Error en versiÃ³n de Padre Hijo]"
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
               descripcion = descripcion + "[No existe aplicaciÃ³n en nodo destino]"
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
               descripcion = descripcion + "[AplicaciÃ³n de origen/destino no existe en tabla SRM]"
          Case "12"
               descripcion = descripcion + "[Tabla de contexto llena (SRM VAX)]"
          Case "13"
               descripcion = descripcion + "[SesiÃ³n invÃ¡lida]"
          Case "14"
               descripcion = descripcion + "[CÃ³digo contexto invÃ¡lido]"
          Case "15"
               descripcion = descripcion + "[AplicaciÃ³n de destino no activa]"
          Case "16"
               descripcion = descripcion + "[No existe PID de proceso originador]"
          Case "17"
               descripcion = descripcion + "[FILE TRANSFER en progeso]"
          Case "18"
               descripcion = descripcion + "[FILE TRANSFER terminado]"
          Case "19"
               descripcion = descripcion + "[CÃ³digo desconocido en campo tipo mensaje]"
          Case "20"
               descripcion = descripcion + "[Mensaje de origen no tiene formato SRM]"
          Case "21"
               descripcion = descripcion + "[Mensaje de respuesta no tiene formato SRM]"
          Case "22"
               descripcion = descripcion + "[Enlace no autorizado a SRM]"
          Case "23"
               descripcion = descripcion + "[SRM inactivo en nodo local]"
          Case "24"
               descripcion = descripcion + "[NODO ALIAS no estÃ¡ en tabla SRM]"
          Case "25"
               descripcion = descripcion + "[Error en largo en data transmitida]"
          Case "26"
               descripcion = descripcion + "[Programa cancelado por problema no determinado]"
          Case "27"
               descripcion = descripcion + "[No hay memoria suficiente]"
          Case "28"
               descripcion = descripcion + "[Se rompiÃ³ enlace]"
          Case "29"
               descripcion = descripcion + "[Nodo origen desconocido]"
          Case "30"
               descripcion = descripcion + "[PID invÃ¡lido]"
          Case "31"
               descripcion = descripcion + "[SRM bloqueado por el operador]"
          Case "32"
               descripcion = descripcion + "[Nodo cerrado]"
          Case "33"
               descripcion = descripcion + "[Nodo origen no existe en tabla SRM]"
          Case "34"
               descripcion = descripcion + "[No hay LU's enviadoras en nodo de destino]"
          Case "35"
               descripcion = descripcion + "[InformaciÃ³n de seguridad inconsistente]"
          Case "36"
               descripcion = descripcion + "[No autorizado]"
          Case "37"
               descripcion = descripcion + "[Error en comando CICS]"
          Case "38"
               descripcion = descripcion + "[FallÃ³ carga o enganche de PGM]"
          Case "39"
               descripcion = descripcion + "[Fuera de horario]"
          Case "40"
               descripcion = descripcion + "[No corresponde]"
          Case "41"
               descripcion = descripcion + "[CondiciÃ³n de excepciÃ³n en manejo de archivos]"
          Case "42"
               descripcion = descripcion + "[Error en HOOK-LINK abortado]"
          Case "61"
               descripcion = descripcion + "[GeneraciÃ³n de PID-OK]"
          Case "62"
               descripcion = descripcion + "[Se validarÃ¡ autorizaciÃ³n]"
          Case "63"
               descripcion = descripcion + "[No se valida autorizaciÃ³n]"
          Case "64"
               descripcion = descripcion + "[Solicita validaciÃ³n de PID]"
          Case "65"
               descripcion = descripcion + "[Seguridad completa OK]"
               
     End Select

     If st <> "00" Then
         descripcion = descripcion + Chr(13)
     End If
     
     Select Case retsrm
         Case "53"
              descripcion = descripcion + "[Vista no existe]"
         Case "54"
              descripcion = descripcion + "[OperaciÃ³n no existe]"
         Case "55"
              descripcion = descripcion + "[Error en formato (clave o datos de input)]"
         Case "56"
              descripcion = descripcion + "[CÃ³digo de transacciÃ³n no numÃ©rica]"
         Case "57"
              descripcion = descripcion + "[CÃ³digo de transacciÃ³n no existe o no inicializada]"
         Case "58"
              descripcion = descripcion + "[MÃ¡ximo de transacciones alcanzado, no se puede comenzar una nueva]"
         Case "59"
              descripcion = descripcion + "[Error al leer archivo de transacciones]"
         Case "60"
              descripcion = descripcion + "[Error al grabar archivo de transacciones]"
         Case "61"
              descripcion = descripcion + "[Error de dato en el Host]"
         Case "65"
              descripcion = descripcion + "[OperaciÃ³n no autorizada para la vista especificada]"
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
              descripcion = descripcion + "[SuperÃ³ intentos permitidos para login]"
         Case "80"
              descripcion = descripcion + "[Registro Bloqueado]"
         Case "81"
              descripcion = descripcion + "[Tabla de locks estÃ¡ llena]"
         Case "82"
              descripcion = descripcion + "[Error en formato de datos]"
         Case "83"
              descripcion = descripcion + "[Registro no estÃ¡ bloqueado]"
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
              descripcion = descripcion + "[Requerimientos invÃ¡lidos (INVREQ)]"
         Case "94"
              descripcion = descripcion + "[IOERR o error al abriri el archivo]"
         Case "95"
              descripcion = descripcion + "[LENGERR]"
         Case "96"
              descripcion = descripcion + "[Dato solicitado no se encontrÃ³]"
         Case "97"
              descripcion = descripcion + "[Archivos cerrados]"
         Case "98"
              descripcion = descripcion + "[NOTSPACE]"
         Case "99"
              descripcion = descripcion + "[BROWSE END]"
     End Select
     descripcion = descripcion + "Llave : [" + ERR_LLAVE + "]" + Chr$(13)
     descripcion = descripcion + "Nodo  : [" + NODO + "]"

     MsgBox descripcion + Chr(13) + "RET: " + Format$(retfun, "00") + "   STATUS: " + st + "   RET_SRM: " + retsrm, 16, "Problemas de ComunicaciÃ³n"
     
     End
End Sub


Function Prepara_Msg()
   Vista = NumVista
   Select Case NumVista
          Case "1828"
               If Environ$("DESA") = "DESA" Then
                  Nodo_Dst = "U005"
               Else
                  Nodo_Dst = "U025"
               End If
               Apl_Dst = "SD20"
               Modo = "L"
               Usuario = "11223344" + Vista + Modo
               LlaveUp = llave
          Case Else
               If Environ$("DESA") = "DESA" Then Nodo_Dst = "UDES"
               Apl_Dst = "SD20"
               Modo = "X0000"
               Usuario = "11223344" + Vista + Modo
               LlaveUp = llave
   End Select
            
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
   Dim itope, Data, continua, tam_reg As String
   Dim Codigo_Prod As String * 3
   Dim i, Ok As Integer
                 
   Codigo_Prod = "XXX"
   Ok = False
   Select Case Manera
      Case 3
           Ok = False
   End Select
   Separa = Ok
End Function
