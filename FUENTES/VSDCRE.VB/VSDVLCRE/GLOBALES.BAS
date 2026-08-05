Attribute VB_Name = "GLOBALES"
Global Version As String * 45

Rem     ==================================================
Rem               Declaraciones de Manejo archivos (.INI)
Rem     ==================================================

Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer


Rem     ===========================================
Rem                 Declaraciones de AIS
Rem     ===========================================
Declare Function AisGetUsr Lib "C:\Bin\F\Wais.Dll" (ByVal Estructura As String) As Integer
Declare Function LInsert Lib "c:\bin\f\waislog.dll" (ByVal nodo As String, ByVal prall As String, ByVal sdall As String, ByVal trall As String, ByVal ctall As String, ByVal flags As String, ByVal autorizo As String, ByVal resultado As String, ByVal perError As String, ByVal sdoError As String, ByVal logdata As String, ByVal Largo As Integer, ByVal fecha_s As String, ByVal hora_S As String, ByVal ind As String, ByVal Funcion As Integer, ByVal msg As String) As Integer


Rem     ===========================================
Rem           Declaraciones de Llamados Shell
Rem     ===========================================

Declare Function FindWindow% Lib "user" (ByVal lpClassName As Any, ByVal lpCaption As Any)

Declare Function GetModuleUsage% Lib "Kernel" (ByVal hProgram%)

