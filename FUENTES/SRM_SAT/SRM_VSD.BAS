Attribute VB_Name = "Srm_Vsd"
' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer
Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long
' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
'Private Declare Function GetModuleUsage Lib "KERNEL32.DLL" (ByVal handler As Long) As Long

Public Const SINNODO = "9999"
Public Const EJECUTABLE = "Srm_Vsd.EXE"
Public Const SINENCONTRAR = "999999"
Public Const SINTRANSACCION = "XXXXXX"

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_INI02 = "ErrorSRM.INI"
Public Const ARCHIVO_MDB01 = "EnvioSRM.MDB"
Public Const ARCHIVO_MDB02 = "Visado.MDB"

Public Const DIRECTORIO_02 = "C:\Bin\S\"
Public Const DIRECTORIO_03 = "C:\Data\V_CtaCte\"
Public Const ARCHIVO_PROFI = "V_CTACTE.EXE"
Public Const ARCHIVO_INI03 = "V_CTACTE.INI"
Public Const DIRECTORIO_SGT = "C:\Data\Admsgt\"

Global Ambiente As String
Global Usuario As String

Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String

lpReturnedString = String$(255, 0)
nSize% = Len(lpReturnedString)
ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
SacaValorINI = Left$(lpReturnedString$, ValorGet)
  
End Function

Public Sub EjecutarPrograma(ByVal Programa As String, ByVal Modo As Integer)
' FI2026 - Retorno era Integer: Shell() en VB6 32-bit devuelve el process ID
' (muy por encima de 32767), que desborda un Integer.
Dim Retorno As Long
  Retorno = Shell(Programa, Modo)
  DoEvents
  ' FI2026 - restaura la espera del original 16-bit:
  '     While GetModuleUsage(Retorno) <> 0 : DoEvents : Wend
  ' EjecutarPrograma es SINCRONA: quien la llama asume que el programa
  ' externo ya termino al volver.
  Call WaitProcess(Retorno)
End Sub

' FI2026 - equivalente Win32 del patron Win16 While GetModuleUsage(pid) <> 0
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

Public Function BuscaNodoOfi(CodOfi As String) As String

Dim BaseSGT As Database
Dim p_Org As Recordset

Set BaseSGT = OpenDatabase(DIRECTORIO_SGT + "TABSGT.MDB")
Set p_Org = BaseSGT.OpenRecordset("Select * From TBLORG Where ORG_ORGCUI=" + CodOfi, dbOpenDynaset)
If p_Org.RecordCount > 0 Then
  BuscaNodoOfi = p_Org!Org_OrgScj
Else
  BuscaNodoOfi = SINNODO
End If
p_Org.Close
BaseSGT.Close


End Function
