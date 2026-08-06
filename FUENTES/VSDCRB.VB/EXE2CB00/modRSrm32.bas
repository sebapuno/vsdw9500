Attribute VB_Name = "modRSrm32"
Option Explicit

' FI2026 - Srmw32 jp
'Declare Function Srmw32 Lib "c:\bin\f\srmw32.dll" (ByVal shost As String, ByVal sapli As String, ByVal smensaje As String, slargo_mensaje As Long, ByVal sstatus As String, ByVal sfuncion As String, ByVal sctx As String, ByVal scontrol As String) As Integer
Declare Function Srmw32 Lib "c:\bin\f\SRMW32.DLL" (ByVal Nodo As String, ByVal Servidor As String, ByVal Mensaje As String, largo As Long, ByVal Status As String, ByVal funcion As String, ByVal Contexto As String, ByVal CONTROLES As String) As Long 'Integer
' Declare Function Srmw32Ex Lib "c:\bin\f\srmw32.DLL" (ByVal Nodo As String, ByVal Servidor As String, ByVal Mensaje As String, largo As Long, ByVal Status As String, ByVal Funcion As String, ByVal Contexto As String, ByVal CONTROLES As String, ByVal AliasPC As String, ByVal NumSeg As Integer, ByVal NumMil As Integer) As Integer
'ant Declare Function Srmw32Ex Lib "c:\bin\f\SRMW32.DLL" (ByVal nodo As String, ByVal Servidor As String, ByVal mensaje As String, largo As Long, ByVal Status As String, ByVal Funcion As String, ByVal Contexto As String, ByVal CONTROLES As String, ByVal AliasPC As String, ByVal NumSeg As Long, ByVal NumMil As Long) As Long
 Declare Function Srmw32Ex Lib "c:\bin\f\SRMW32.DLL" (ByVal Nodo As String, ByVal Servidor As String, ByVal Mensaje As String, largo As Long, ByVal Status As String, ByVal funcion As String, ByVal Contexto As String, ByVal CONTROLES As String, ByVal AliasPC As String, ByVal NumSeg As Integer, ByVal NumMil As Integer) As Integer


'jp Declare Function Srmw32ex Lib "c:\bin\f\srmw32.DLL" (ByVal Nodo As String, ByVal Servidor As String, ByVal Mensaje As String, largo As Long, ByVal Status As String, ByVal funcion As String, ByVal Contexto As String, ByVal CONTROLES As String, ByVal AliasPC As String, ByVal NumSeg As Integer, ByVal NumMil As Integer) As Integer

Type Type_ParamSrm8K
    Nodo            As String * 5       ' Nodo Servidor.
    Servidor        As String * 5       ' Aplicacion Servidora.
    Mensaje         As String * 8200    ' Mensaje.
    largo           As Long             ' Largo del Mensaje. FI2026: era Integer (16-bit);
                                         ' el DLL espera 'int*largo' (32-bit) por referencia.
                                         ' Con Integer + CLng() en la llamada, el ByRef se
                                         ' rompe (VB crea un temporal Long) y el largo real
                                         ' de la RESPUESTA nunca vuelve a este campo.
    Status          As String * 3       ' Status.
    funcion         As String * 3       ' Funcion.
    Contexto        As String * 3       ' Contexto.
    Control         As String * 9000    ' Control.
    base            As String           ' Base    Sybase.-
    Usuario         As String           ' Usuario Sybase.-
    APartirDe       As Integer          ' A partir de que registro.-
    ExisteSINO      As String           ' Existencia de mas registros S o N.
    retorno         As Integer          ' Retorno llamada SRM
    RetServidor     As Integer          ' Retorno Servidor de Datos
End Type
Global ParamSrm8K   As Type_ParamSrm8K
Private Const HEADERSY = "b12256781234W"

'************************************************************************************
' Funciones AISWindows
'************************************************************************************
'Declare Function AISGetUsr Lib "c:\bin\f\wais32.dll" (ByVal Datos As String) As Integer

Private Const SISTEMA = "SMP"
Private Const Aplicacion = "SMP_BACK_OFFICE"
Dim MODOLOGUEO As Integer
  ' 1 es remoto
  ' 2 es local
  ' 3 es buffering

Global codtran As String, destran As String, rutcli As String, rutsup As String, datcon As String
' FI_MIG - GetModuleUsage no existe en Kernel32 (es API Win16). Reemplazado por WaitProcess.
'Declare Function GetModuleUsage% Lib "Kernel32" (ByVal hProgram%)
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwAccess As Long, ByVal fInherit As Long, ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Private Const VERSION_APP = "v3.0"

' FI2026: sBdd/sApliOri/sNodoOri se usaban en SendQuery (bloque "IBM
' MIGRACION SYBASE") sin declarar en ningun .bas de WCRE9500 ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬ï¿½ bug
' preexistente que nunca compilo (mismo caso que Valida_Migracion, ver
' modAcceso.bas). Se declaran para que compile; sApliOri/sNodoOri nunca
' se asignan en el original tampoco, asi que quedan "" (mismo
' comportamiento, no se le agrega logica nueva).
Dim sBdd As String
Dim sApliOri As String
Dim sNodoOri As String

Global NOMBRE_USUARIO As String
Global RUT_USUARIO As String
Global DV_USUARIO  As String

Global NAME_BD As String
Global SERV_BD As String
Global NODO_BD As String
Global LOGOINI As String

Dim HABIL_APLI As String
Dim HABIL_NODO As String
Dim HABIL_TIPOPE As String
Dim HABIL_VISTA As String
Dim HABIL_TABLA As String

'FI2025 SHG -- Se incorporan variables globales para almacenar informaciÃƒÆ’Ã‚Â³n de RUT y DV
Global xDVusr As Double
Global xRUTusr As Double

'---------------------------------------------
' LIMPIA_CADENA
'
' Saca toda la parte que no sirve en la data
' que devuelve el srm.
'---------------------------------------------
Function Limpia_Cadena(Lc_String As String) As String
Dim Li_Cont         As Integer
Dim i               As Integer
Dim Lc_Limpia_Data  As String

Li_Cont = 0
For i = Len(Lc_String$) To 1 Step -1
    If Mid$(Lc_String$, i, 1) = "~" Then
        Li_Cont = i
        Lc_Limpia_Data$ = Mid$(Lc_String$, 1, Li_Cont)
        Exit For
    End If
Next i

If Li_Cont = 0 Then
    Lc_Limpia_Data$ = Trim$(Lc_String$)
End If

Limpia_Cadena = Lc_Limpia_Data$

End Function

' FI2026 SHG - num_srm comentada: GetSceIni sin implementacion
'Function num_srm(numero As Variant) As String
' Dim Lc_C As String
' Dim i As Integer
' Dim campo As String
' Dim Dec As String, Miles As String
' 
' Dec = GetSceIni("Intl", "sDecimal", "win.ini")
' Miles = GetSceIni("Intl", "sThousand", "win.ini")
' 
' campo = CStr(numero)
' Lc_C = ""
' For i = 1 To Len(Trim$(campo))
'     If Mid$(campo, i, 1) = Dec Then
'         Lc_C = Lc_C & "."
' 
'     ElseIf Mid$(campo, i, 1) <> Miles Then
'         Lc_C = Lc_C & Mid$(campo, i, 1)
'       
'     End If
' Next
' 
' num_srm = Lc_C
' 
'End Function

Function Retornos_Srm(Pi_RetSrm As Integer) As String

Select Case Pi_RetSrm
       Case 0:
            Retornos_Srm = "00 Operacion Exitosa"
       Case 1:
            Retornos_Srm = "01 Error FBDD SRPI"
       Case 2:
            Retornos_Srm = "02 Proceso abortado por otra interrupciÃƒÆ’Ã‚Â³n"
       Case 3:
            Retornos_Srm = "03 Fallo obtencion de variable de ambiente LAN"
       Case 4:
            Retornos_Srm = "04 Fallo apertura de log file"
       Case 5:
            Retornos_Srm = "05 Fallo inicializacion de comunicaciones"
       Case 6:
            Retornos_Srm = "06 Mensaje aceptado, habia respuesta previa"
       Case 7:
            Retornos_Srm = "07 Envio en ejecucion. Esperando respuesta"
       Case 9:
            Retornos_Srm = "09 FunciÃƒÆ’Ã‚Â³n incorrecta en llamado padre hijo"
       Case 10:
            Retornos_Srm = "10 Fallo obtencion de 'LUNUMBER'"
       Case 11:
            Retornos_Srm = "11 Fallo obtencion de 'TIMEOUT'"
       Case 21:
            Retornos_Srm = "21 Fallo obtencion de 'CICS_NAME'"
       Case 30:
            Retornos_Srm = "30 Fallo obtencion de 'NODENAME'"
       Case 35:
            Retornos_Srm = "35 Error Memoria"
       Case 41:
            Retornos_Srm = "41 Error en apertura de canal de comunicaciÃƒÆ’Ã‚Â³n"
       Case 42:
            Retornos_Srm = "42 Error de escritura al canal de comunicaciÃƒÆ’Ã‚Â³n"
       Case 43:
            Retornos_Srm = "43 Error de lectura del canal de comunicaciÃƒÆ’Ã‚Â³n"
       Case 44:
            Retornos_Srm = "44 Error de cierre del canal de comunicaciÃƒÆ’Ã‚Â³n"
       Case 45:
            Retornos_Srm = "45 Error en largo de la data de respuesta"
       Case 90:
            Retornos_Srm = "90 Error en inicializaciÃƒÆ’Ã‚Â³n de comunicaciones"
       Case 91:
            Retornos_Srm = "91 Error en inicializaciÃƒÆ’Ã‚Â³n de comunicaciones"
       Case 92:
            Retornos_Srm = "92 Fallaron 3 reintentos"
       Case 93:
            Retornos_Srm = "93 Error en data SRM recibida"
       Case 97:
            Retornos_Srm = "97 Error en carga de proceso externo"
       Case 98:
            Retornos_Srm = "98 Error en contexto de comunicaciÃƒÆ’Ã‚Â³n padre hijo"
       Case 99:
            Retornos_Srm = "99 Error en versiÃƒÆ’Ã‚Â³n de padre / hijo"
       Case Else
            Retornos_Srm = ""
End Select

End Function

Function Retornos_Status() As String

Select Case Trim$(ParamSrm8K.Status)
       Case Is = "00"
            Retornos_Status = "00 OperaciÃƒÆ’Ã‚Â³n Exitosa"
       Case Is = "01"
            Retornos_Status = "01 Servidor de Comunicaciones inactivo en nodo destino"
       Case Is = "02"
            Retornos_Status = "02 LU no definida como receptora en nodo destino"
       Case Is = "03"
            Retornos_Status = "03 No existe aplicaciÃƒÆ’Ã‚Â³n en nodo destino"
       Case Is = "04"
            Retornos_Status = "04 LU no definida en nodo"
       Case Is = "05"
            Retornos_Status = "05 No existe programa asociado en nodo destino"
       Case Is = "06"
            Retornos_Status = "06 Disponible"
       Case Is = "07"
            Retornos_Status = "07 No hay LU disponible en nodo de origen"
       Case Is = "08"
            Retornos_Status = "08 No hay LU disponible en nodo re-ruteador"
       Case Is = "09"
            Retornos_Status = "09 Timeout. aplicaciÃƒÆ’Ã‚Â³n remota no contesta a tiempo"
       Case Is = "10"
            Retornos_Status = "10 Nodo de destino no existe en tabla SRM"
       Case Is = "11"
            Retornos_Status = "11 aplicaciÃƒÆ’Ã‚Â³n de origen/destino no existe en table SRM"
       Case Is = "12"
            Retornos_Status = "12 Tabla de contextos llena"
       Case Is = "13"
            Retornos_Status = "13 SesiÃƒÆ’Ã‚Â³n invalida. SRM no puede mantener SesiÃƒÆ’Ã‚Â³n"
       Case Is = "14"
            Retornos_Status = "14 CÃƒÆ’Ã‚Â³digo contexto invalido"
       Case Is = "15"
            Retornos_Status = "15 aplicaciÃƒÆ’Ã‚Â³n de destino no esta activa"
       Case Is = "16"
            Retornos_Status = "16 No existe PID de proceso originador"
       Case Is = "17"
            Retornos_Status = "17 File Transfer en Progreso"
       Case Is = "18"
            Retornos_Status = "18 File Transfer Terminado"
       Case Is = "19"
            Retornos_Status = "19 CÃƒÆ’Ã‚Â³digo desconocido en campo tipo mensaje"
       Case Is = "20"
            Retornos_Status = "20 Mensaje de origen no tiene formato SRM"
       Case Is = "21"
            Retornos_Status = "21 Mensaje de respuesta no tiene formato SRM"
       Case Is = "22"
            Retornos_Status = "22 Enlace no autorizado a SRM"
       Case Is = "23"
            Retornos_Status = "23 SRM inactivo en nodo local"
       Case Is = "24"
            Retornos_Status = "24 Nodo-Alias no esta en tabla SRM"
       Case Is = "25"
            Retornos_Status = "25 Error de largo en data transmitida"
       Case Is = "26"
            Retornos_Status = "26 Programa cancelado por problema no determinado"
       Case Is = "27"
            Retornos_Status = "27 No hay memoria suficiente"
       Case Is = "28"
            Retornos_Status = "28 aplicaciÃƒÆ’Ã‚Â³n remota no contestara ( Se rompe el enlace )"
       Case Is = "29"
            Retornos_Status = "29 Disponible"
       Case Is = "30"
            Retornos_Status = "30 PID Invalido"
       Case Is = "31"
            Retornos_Status = "31 Disponible"
       Case Is = "32"
            Retornos_Status = "32 Nodo cerrado"
       Case Is = "33"
            Retornos_Status = "33 Nodo de origen no existe en tabla SRM"
       Case Is = "34"
            Retornos_Status = "34 No hay LU's enviadoras para nodo de destino"
       Case Is = "35"
            Retornos_Status = "35 Disponible"
       Case Is = "36"
            Retornos_Status = "36 No autorizado"
       Case Is = "37"
            Retornos_Status = "37 Error en comando CICS"
       Case Is = "38"
            Retornos_Status = "38 Fallo carga o enganche de pgm. Solicitado"
       Case Is = "39"
            Retornos_Status = "39 Fuera de horario"
       Case Is = "40"
            Retornos_Status = "40 No corresponde"
       Case Is = "41"
            Retornos_Status = "41 CondiciÃƒÆ’Ã‚Â³n de exepciÃƒÆ’Ã‚Â³n en manejo de archivos"
       Case Is = "61"
            Retornos_Status = "61 GeneraciÃƒÆ’Ã‚Â³n de PID Ok"
       Case Is = "62"
            Retornos_Status = "62 ValidaciÃƒÆ’Ã‚Â³n de PID Ok ( Se validara autorizaciÃƒÆ’Ã‚Â³n )"
       Case Is = "63"
            Retornos_Status = "63 ValidaciÃƒÆ’Ã‚Â³n de PID Ok ( No se valida autorizaciÃƒÆ’Ã‚Â³n )"
       Case Else
            Retornos_Status = ""
End Select


End Function
'-----------------------------------------------------------------------
' SENDQUERY1
'Este sendquery1 es igual que el sendquery, pero en vez de la llamada
'srmw, se ocupa la llamada srmw8ex para extender el tiempo de respuesta.
'
' Envia Sentencia SQL via SRM y recibe Respuesta.
'
' Parametros :
'       De Input :
'       Gpc_BaseDatos       : Nombre de la Base de Datos.
'       Gpc_Pgm_Servidor    : Programa servidor que tomarÃƒÆ’Ã‚Â¡ la consulta.
'       Gpc_Nodo_Servidor   : Nodo Servidor donde esta localizado el
'                             Pgm_Servidor.
'       Gpc_Sentencia       : Sentencia SQL.
'
'       De Output :
'       s                   : Mensaje de Retorno.
'                             -1 : Error en la base de datos.
'                             de lo contrario retorna la data de salida.
'-----------------------------------------------------------------------
Function SendQuery(Gpc_BaseDatos As String, Gpc_Pgm_Servidor As String, Gpc_Nodo_Servidor As String, Gpc_Sentencia As String) As String
    Dim Lc_BaseDatos    As String
    Dim lc_mensaje      As String
    Dim m               As String
    Dim X               As Integer
    Dim s               As String
    Dim r               As String
    Dim n               As Double
    Dim Lc_Resp_Srm     As String
    Dim Lc_Resp_Sta     As String
    Dim Lc_Resp_Msg     As String
    
    ParamSrm8K.base$ = Gpc_BaseDatos$
    ParamSrm8K.Servidor$ = Gpc_Pgm_Servidor$
    ParamSrm8K.Nodo$ = Gpc_Nodo_Servidor$
        
    Gpc_Sentencia$ = Trim$(Gpc_Sentencia$)
'-------------------------------------------------------------------------
'Saca el largo del nombre de la base de datos - 10 y adiciona los espacios
'faltantes.
'-------------------------------------------------------------------------
    Lc_BaseDatos$ = Trim$(LCase$(ParamSrm8K.base$))
    Lc_BaseDatos$ = Lc_BaseDatos$ + Space$(10 - Len(Lc_BaseDatos$))
'----------------------------------------------------
'Verifica a partir de que registro comienza el Query.
'----------------------------------------------------
    If ParamSrm8K.APartirDe% = 0 Then
       ParamSrm8K.APartirDe% = 1
    End If
'-------------------------------------------------------------------
' Concatena Encabezado fijo + Registro a que comienza el Query +
' "N" + Nombre de la Base de Datos + Sentencia SQL.
'-------------------------------------------------------------------
    lc_mensaje$ = HEADERSY & Format$(ParamSrm8K.APartirDe, "0000000") & "N" & Lc_BaseDatos & Gpc_Sentencia
    ParamSrm8K.Mensaje$ = lc_mensaje$
    ParamSrm8K.largo = Len(lc_mensaje$)
    ParamSrm8K.Status$ = "00"
    ParamSrm8K.funcion$ = "01"
    ParamSrm8K.Contexto$ = "00"
    ParamSrm8K.Control$ = ""
'-------------------
' Envia Mensaje.
'-------------------
'********************************************************************************************
'IBM MIGRACIÃƒÆ’Ã¢â‚¬Å“N SYBASE - INICIO
'OCTUBRE 2010

    'Rescatar Base de Datos
    sBdd = Trim(Mid(ParamSrm8K.Mensaje$, 22, 10))
    'Imprimir Entrada de Datos en Archivo Log
    'Call Genera_log(sNomArchivoLog, "Entrada", ParamSrm8K.Nodo$, ParamSrm8K.Servidor$, ParamSrm8K.Mensaje$, CLng(ParamSrm8K.largo%), ParamSrm8K.Status$, ParamSrm8K.Funcion$, ParamSrm8K.Contexto$, ParamSrm8K.Control$)
    'Validar Datos MigraciÃƒÆ’Ã‚Â³n
    Call Valida_Migracion(ParamSrm8K.Mensaje$, ParamSrm8K.Servidor$, ParamSrm8K.Nodo$)
       
     X% = Srmw32(ParamSrm8K.Nodo$, ParamSrm8K.Servidor$, ParamSrm8K.Mensaje$, ParamSrm8K.largo, ParamSrm8K.Status$, ParamSrm8K.funcion$, ParamSrm8K.Contexto$, ParamSrm8K.Control$) 'Llamada al SRM Original -- FI2026: sin CLng(), asi el DLL SI actualiza el largo real de la respuesta por referencia
    
    'Imprimir Salida de Datos en Archivo Log
    'Call Genera_log(sNomArchivoLog, "Salida", ParamSrm8K.Nodo$, ParamSrm8K.Servidor$, ParamSrm8K.Mensaje$, CLng(ParamSrm8K.largo%), ParamSrm8K.Status$, ParamSrm8K.Funcion$, ParamSrm8K.Contexto$, ParamSrm8K.Control$)

'IBM MIGRACIÃƒÆ’Ã¢â‚¬Å“N SYBASE - TERMINO
'********************************************************************************************
    
    ParamSrm8K.RetServidor% = 0
    ParamSrm8K.retorno% = X%
    Lc_Resp_Srm$ = Retornos_Srm(X%)
    Lc_Resp_Sta$ = Retornos_Status()
        
     ''   MsgBox "ParamSrm8K.retorno% " & ParamSrm8K.retorno%
     ''   MsgBox "Lc_Resp_Srm$ " & Lc_Resp_Srm$
     ''   MsgBox "Lc_Resp_Sta$ " & Lc_Resp_Sta$
        
    If Not (X% = 0 And Left$(ParamSrm8K.Mensaje, 2) = "00") Then
       Lc_Resp_Msg = "Error en Mensaje de Retorno SRM :" & Chr$(13) & Chr$(9)
       Lc_Resp_Msg = Lc_Resp_Msg & Mid$(ParamSrm8K.Mensaje, 1, 31) & Chr$(13) & Chr$(13)
       Lc_Resp_Msg = Lc_Resp_Msg & "Error de Retorno SRM :" & Chr$(13) & Chr$(9)
       Lc_Resp_Msg = Lc_Resp_Msg & Trim$(Lc_Resp_Srm$) & Chr$(13) & Chr$(13)
       Lc_Resp_Msg = Lc_Resp_Msg & "Error de Retorno STATUS :" & Chr$(13) & Chr$(9)
       Lc_Resp_Msg = Lc_Resp_Msg & Trim$(Lc_Resp_Sta$)
       MsgBox Lc_Resp_Msg, 64, "Error SRM"
       
       s$ = "-1"
    Else
        ' FI2026: Mensaje es String*8200 (largo fijo); el DLL solo escribe los
        ' primeros ParamSrm8K.largo bytes de la RESPUESTA y no toca el resto
        ' del buffer (asi tambien se comporta la DLL oficial, verificado con
        ' harness). Leer el buffer completo con Trim$ arrastraba la cola del
        ' REQUEST anterior que quedaba mas alla de la respuesta real.
        If ParamSrm8K.largo > 0 And ParamSrm8K.largo <= Len(ParamSrm8K.Mensaje) Then
            r$ = Trim$(Left$(ParamSrm8K.Mensaje, ParamSrm8K.largo))
        Else
            r$ = Trim$(ParamSrm8K.Mensaje)
        End If
        n# = Val(Mid$(r$, 4, 7))
        If n# > 0 Then
           s$ = Right$(r$, Len(r$) - 14)
'----------------------------------------------------------------------------
'Limpia mensaje de retorno del Srm, dejando solo los campos a ser utilizados.
'----------------------------------------------------------------------------
           s$ = Limpia_Cadena(s$)
        Else
            s$ = ""
        End If
    End If
    
    If Mid$(r$, 3, 1) = "S" Then
        ParamSrm8K.APartirDe = ParamSrm8K.APartirDe + n#
        '+ 1
        ParamSrm8K.ExisteSINO = "S"
    Else
        ParamSrm8K.APartirDe = 1
        ParamSrm8K.ExisteSINO = "N"
    End If
    
    SendQuery = s$
    
'********************************************************************************************
'IBM MIGRACIÃƒÆ’Ã¢â‚¬Å“N SYBASE - INICIO
'OCTUBRE 2010
    'Asigno Valores de origen a las variables
    ParamSrm8K.Servidor$ = sApliOri   'SERVIDOR
    ParamSrm8K.Nodo$ = sNodoOri  'NODO
    
'IBM MIGRACIÃƒÆ’Ã¢â‚¬Å“N SYBASE - TERMINO
'********************************************************************************************
    
    
End Function



' FI_MIG - Reemplaza el patron Win16 Do While GetModuleUsage%(pid) Loop
' Espera a que el proceso pid termine antes de continuar.
Public Sub WaitProcess(ByVal pid As Long)
    Dim hProc As Long
    Const SYNCHRONIZE As Long = &H100000
    hProc = OpenProcess(SYNCHRONIZE, 0, pid)
    If hProc <> 0 Then
        Do While WaitForSingleObject(hProc, 100) <> 0
            DoEvents
        Loop
        CloseHandle hProc
    End If
End Sub

' FI2026 SHG - srm_num comentada: GetSceIni sin implementacion
'Function srm_num(campo As String) As Variant
' Dim Lc_C As String
' Dim i As Integer, Dec As String
' 
' Dec = GetSceIni("Intl", "sDecimal", "win.ini")
' 
' Lc_C = ""
' For i = 1 To Len(Trim$(campo))
'     If Mid$(campo, i, 1) = "." Then
'         Lc_C = Lc_C & Dec
'     Else
'         Lc_C = Lc_C & Mid$(campo, i, 1)
'     End If
' Next
' 
' srm_num = Lc_C
' 
'End Function

'--------------------------------------------------------------------
' TRAE_CAMPO
'
' Separa campo de la cadena recibida a travÃƒÆ’Ã‚Â©s
' del SRM.
'
' Parametros :
'   Data_Resp = String con mensaje recibido del SRM.
'   Lgi_Largo = PosiciÃƒÆ’Ã‚Â³n del caracter inicial dentro del string.
'--------------------------------------------------------------------
Function Trae_Campo(Data_Resp As String, Lgi_Largo As Integer) As String
    Dim Lgc_Campo As String
    Dim i As Integer

    Lgc_Campo$ = ""
    For i = Lgi_Largo To Len(Data_Resp$)
       If Mid$(Data_Resp$, i, 1) = "~" Then
          Exit For
       End If
       Lgc_Campo$ = Lgc_Campo$ & Mid$(Data_Resp$, i, 1)
    Next i
    Trae_Campo = Lgc_Campo$
End Function
