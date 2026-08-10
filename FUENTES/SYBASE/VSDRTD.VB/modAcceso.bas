Attribute VB_Name = "modAcceso"
Option Explicit

Global FUNCIONALIDAD As Integer
Global TIPO_RUT As Integer  '0 => utilice rut AIS
                            '1 => utilice rut ingresado por usuario
Global RET_ACCESO As Integer  'true usuario ingresó clave correcta
                              'false usuario ingresó clave incorrecta


' FI2026: stub de Valida_Migracion -- modRSrm32.bas la llama (Call
' Valida_Migracion(...) en SendQuery) pero no existe en ningun .bas de
' WCRE9500 (ni siquiera comentado con el nombre correcto: el stub original
' aqui decia "Valida_Migracion_ant", con sufijo "_ant" que no matchea la
' llamada real -- confirmado que esto no compila ni en el WCRE9500 original,
' es un bug preexistente nunca detectado porque la compilacion se bloquea
' antes por la licencia de THREED32.OCX). Sin efecto -- solo satisface la
' referencia para que el proyecto compile.
Public Sub Valida_Migracion(p_sBaseApli As String, p_sApli As String, p_sNodoApli As String)
End Sub
Function Usuario_Valido(pr_funcion As String) As Boolean ' no se usa
Dim DatosUser   As String * 200
Dim nRet        As Integer

Dim sw_requiere_clave As Boolean
Dim Li_Largo As Integer
Dim sqlq As String, r As String
Dim aux_funcion As String, aux_monope As String
Dim aux_monmax As String, aux_reqcla As String
Dim aux_nombre As String, aux_prefijo As String
Dim li_pos1, li_pos2 As Integer

'inicialización de Variables
sw_requiere_clave = False
Usuario_Valido = False

'Valor por defecto de la funcionalidad (se usará si el SP no devuelve aux_funcion numérico)
If IsNumeric(pr_funcion) Then
   FUNCIONALIDAD = CInt(pr_funcion)
End If


'---- HABILITAR ESTAS LINEAS SI HAY PROBLEMAS DE USUARIO EN EL AIS ---
'RUT_USUARIO = 7572376
'DV_USUARIO = "8"
'NOMBRE_USUARIO = "Ricardo Estay"
'---- HABILITAR ESTAS LINEAS SI HAY PROBLEMAS DE USUARIO EN EL AIS ---

'sinACC - Control de acceso deshabilitado
'TIPO_RUT = 1
'ctrclave.Show 1
'If RET_ACCESO Then ...
RUT_USUARIO = 0
DV_USUARIO = "0"
NOMBRE_USUARIO = "DESARROLLO"
RET_ACCESO = True
  

  
' ## SE COMENTA CODIGO ORIGINAL ##
                'Verifica si el usuario está conectado al Ais Windows
                'nRet = AISGetUsr(DatosUser)
                ''If nRet = 0 Then
                '   MsgBox "Usted no está conectado en AISWindows", 16, "Control de Acceso"
                '   Exit Function
                'Else


                ' ----------------- EJEMPLO-----------
                ''   DV_USUARIO = Mid$(DatosUser, 9, 1)
                ''   RUT_USUARIO = Mid$(DatosUser, 1, 8)
                ''   RUT_USUARIO = Trim$(Str$(Val(RUT_USUARIO)))
                ''   li_pos1 = InStr(1, DatosUser, "_", 1)
                ''   If li_pos1 > 0 Then
                ''      li_pos2 = InStr(li_pos1 + 1, DatosUser, "_", 1)
                ''      If li_pos2 > 0 Then
                ''         NOMBRE_USUARIO = Mid(DatosUser, li_pos1 + 1, (li_pos2 - 1) - (li_pos1))
                
' FI2025 SHG
    RUT_USUARIO = xRUTusr
    DV_USUARIO = xDVusr
   RUT_USUARIO = Trim$(Str$(Val(RUT_USUARIO)))
   li_pos1 = InStr(1, xRUTusr, "_", 1)
   If li_pos1 > 0 Then
      li_pos2 = InStr(li_pos1 + 1, DatosUser, "_", 1)
      If li_pos2 > 0 Then
         NOMBRE_USUARIO = Mid(DatosUser, li_pos1 + 1, (li_pos2 - 1) - (li_pos1))
      End If
   End If
'End If

'SQL
sqlq = "exec pro_smp0080 1," & pr_funcion & "," & RUT_USUARIO

'Llamada SRM
ParamSrm8K.APartirDe = 1
r = SendQuery(NAME_BD, SERV_BD, NODO_BD, sqlq)

'Validaciones del retorno
If r = "-1" Then
   Exit Function
End If

If r = "" Then
   MsgBox "Usted no est" & Chr$(225) & " autorizado para utilizar esta aplicaci" & Chr$(243) & "n", vbCritical, "Control Acceso" 'jp2
   Exit Function
End If

Li_Largo% = 1
'Lee datos
aux_funcion = Trae_Campo(r, Li_Largo%)
Li_Largo% = Li_Largo% + Len(aux_funcion) + 1
aux_monope = Trae_Campo(r, Li_Largo%)
Li_Largo% = Li_Largo% + Len(aux_monope) + 1
aux_monmax = Trae_Campo(r, Li_Largo%)
Li_Largo% = Li_Largo% + Len(aux_monmax) + 1
aux_reqcla = Trae_Campo(r, Li_Largo%)
Li_Largo% = Li_Largo% + Len(aux_reqcla) + 1
aux_nombre = Trae_Campo(r, Li_Largo%)
Li_Largo% = Li_Largo% + Len(aux_nombre) + 1
aux_prefijo = Trae_Campo(r, Li_Largo%)
Li_Largo% = Li_Largo% + Len(aux_prefijo) + 1

'Verificar si requiere clave
If Trim(aux_reqcla) <> "" Then
   sw_requiere_clave = True
   FUNCIONALIDAD = CInt(aux_funcion)
End If
     
'If sw_requiere_clave = True Then
'   TIPO_RUT = 1 '0
'   frmAcceso.Show 1
'   If RET_ACCESO Then
'      Usuario_Valido = True
'   Else
'      Usuario_Valido = False
'   End If
'Else
   Usuario_Valido = True
'End If

    
End Function

Function encripta(tex_clave As String) As String
Dim cl As String, i As Integer, a As Integer, b As Integer
Dim Num As Long

  cl = ""
  For i = 1 To Len(tex_clave)
    Num = (i + Len(tex_clave)) * Asc(Mid$(tex_clave, i, 1))
    a = Int(Val(Num / 93))
    b = Num - a * 93 + 32
    If b = 39 Or b = 34 Or b = 126 Then
       b = b - 1
    End If
    cl = cl + Chr$(b)
  Next i
  encripta = cl

End Function

' --- FI2026 - Inicializar_Acceso: punto de entrada estandar desde Form_Load
' Lee parametros del INI y valida el usuario. Uso tipico en Form_Load:
'   If Not Inicializar_Acceso(DIRECTORIO_01 & ARCHIVO_INI01) Then End
'   IdUsuario = Format$(Val(RUT_USUARIO), "00000000")
' Requiere: funcion SacaValorINI disponible en el proyecto
' Secciones requeridas en el INI:
'   [BASEACCESO]   NODO=, APLI=, BASEACC=
'   [FUNCIONALIDAD]  FUNCIONALIDAD=
' FI2026 SHG - Inicializar_Acceso comentada: SacaValorDeINI sin implementacion
' Function Inicializar_Acceso(sPathINI As String) As Boolean
'     NAME_BD = SacaValorDeINI("BASEACCESO", "BASEACC", sPathINI)
'     SERV_BD = SacaValorDeINI("BASEACCESO", "APLI", sPathINI)
'     NODO_BD = SacaValorDeINI("BASEACCESO", "NODO", sPathINI)
'     FUNCIONALIDAD = CInt(SacaValorDeINI("FUNCIONALIDAD", "FUNCIONALIDAD", sPathINI))
'   'jp2  Inicializar_Acceso = Usuario_Valido(CStr(FUNCIONALIDAD))
' End Function

' --- FI2026 - DvCalc: calcula digito verificador de RUT
' Implementado inline (CREDITOS no tiene ObtenerDigitoVerificador)
Function DvCalc(t As String) As String
Dim i As Integer, suma As Integer, mult As Integer, resto As Integer
Dim Num As Long
    Num = Val(t)
    suma = 0: mult = 2
    Do While Num > 0
        suma = suma + (Num Mod 10) * mult
        Num = Num \ 10
        mult = mult + 1
        If mult > 7 Then mult = 2
    Loop
    resto = 11 - (suma Mod 11)
    If resto = 11 Then
        DvCalc = "0"
    ElseIf resto = 10 Then
        DvCalc = "K"
    Else
        DvCalc = CStr(resto)
    End If
End Function


