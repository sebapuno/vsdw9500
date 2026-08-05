Attribute VB_Name = "Módulo1"
Option Explicit

Declare Function AisGetUsr Lib "C:\Bin\F\Wais.Dll" (ByVal Estructura As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const DIRECTORIO_02 = "C:\Data\AdmSgt\"

Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSrm.MDB"
Public Const ARCHIVO_MDB03 = "TABSGT.MDB"

Public Const RUTDEFECTO = "12345678"

Global BaseVisado As Database
Global BaseEnvioSrm As Database

