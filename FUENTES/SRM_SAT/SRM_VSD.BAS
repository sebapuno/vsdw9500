Attribute VB_Name = "Srm_Vsd"
Declare Function AisGetUsr Lib "C:\Bin\F\Wais.Dll" (ByVal Estructura As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer
Private Declare Function GetModuleUsage Lib "kernel.dll" (ByVal handler As Integer) As Integer

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
Dim Retorno As Integer
  Retorno = Shell(Programa, Modo)
  DoEvents
  While GetModuleUsage(Retorno) <> 0
     DoEvents
  Wend
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
