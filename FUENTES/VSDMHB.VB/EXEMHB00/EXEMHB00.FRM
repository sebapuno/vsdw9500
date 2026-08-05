VERSION 4.00
Begin VB.Form frmPrincipal 
   Caption         =   "ExeCRE00"
   ClientHeight    =   1605
   ClientLeft      =   4260
   ClientTop       =   1755
   ClientWidth     =   3570
   Height          =   2010
   Left            =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   3570
   Top             =   1410
   Visible         =   0   'False
   Width           =   3690
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim IdUsuario As String * 8
Dim DataINI As String * 800
Dim strEstado As String * 1
Dim Vista As String
Dim Operacion As String * 1

Private Sub Form_Load()
Dim PunteroMsg As Recordset
Dim PunteroVsdPrd As Recordset
Dim RutCliente As String
Dim Producto As String
Dim Correl As Integer
Dim P As Boolean
Dim Data1, Data2 As String

If Not App.PrevInstance Then
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set BaseEnvioSrm = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
   
  RutCliente = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Correl = Val(SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01))
 
  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
  PunteroVsdPrd.Seek "=", RutCliente, Producto, Correl
  DataINI = Trim$(PunteroVsdPrd!Prd_Data)
  PunteroVsdPrd.Close
   
  strEstado = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)
  
  If strEstado = "1" Then
    strEstado = "A"
  Else
    strEstado = "R"
  End If
  
  Data1 = "b12256781234W00000012plt_emp01 exec modifica_estado_macb "
  
  Data2 = Mid$(Trim$(RutCliente), 1, Len(Trim$(RutCliente)) - 1) + ","
  Data2 = Data2 + "'" + Mid$(DataINI, 47, 2) + "/" + Mid$(DataINI, 49, 2) + "/" + Mid$(DataINI, 43, 4) + "'"
  Data2 = Data2 + ",'" + strEstado + "'"
  Call ActualizarEnvioSRM(Data1, Data2)
        
  Data1 = "b12256781234W00000012plt_emp01 exec elimina_tabla_macb "
  
  Data2 = "'" + Mid$(DataINI, 1, 3) + "',"
  Data2 = Data2 + "'" + Mid$(DataINI, 4, 3) + "',"
  Data2 = Data2 + "'" + Mid$(DataINI, 7, 3) + "',"
  Data2 = Data2 + "'" + Mid$(DataINI, 40, 1) + "',"
  Data2 = Data2 + "'" + Mid$(DataINI, 41, 2) + "',"
  Data2 = Data2 + "'" + Mid$(DataINI, 43, 8) + "',"
  Data2 = Data2 + "'" + Mid$(DataINI, 51, 10) + "',"
  Data2 = Data2 + "'" + Mid$(DataINI, 337, 8) + "' "
  Call ActualizarEnvioSRM(Data1, Data2)
  
End If

BaseEnvioSrm.Close
BaseVisado.Close

P = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")

End

End Sub
