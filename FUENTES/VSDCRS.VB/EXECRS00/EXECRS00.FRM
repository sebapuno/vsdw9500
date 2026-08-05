VERSION 4.00
Begin VB.Form frmPrincipal 
   Caption         =   "ExeCRE00"
   ClientHeight    =   1800
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   7335
   Height          =   2205
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   7335
   Top             =   1170
   Width           =   7455
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
Dim OperCRE As String
Dim DataPLT As String
Dim Data1 As String
Dim Data2 As String
Dim Data3 As String
Dim NomCliente As String
Dim CodEjecutivoVenta As String
Dim RutEjecutivo As String
Dim i As Integer
Dim RutAsegurado As String
Dim Encabezado As String
Dim Llaveproducto As String
Dim OperCRS As String
Dim VarSiebel As String

If Not App.PrevInstance Then
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set BaseEnvioSrm = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
   
  RutCliente = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Correl = Val(SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01))
 
  VarSiebel = SacaValorINI("SYBASE", "SIEBEL", DIRECTORIO_01 + ARCHIVO_INI01)
 
  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
  PunteroVsdPrd.Seek "=", RutCliente, Producto, Correl
  DataINI = PunteroVsdPrd!Prd_Data
  PunteroVsdPrd.Close
   
  OperCRS = Mid$(DataINI, 40, 1)
   
  Encabezado = Mid$(DataINI, 1, 6)
  Llaveproducto = Mid$(DataINI, 10, 30)
  
  strEstado = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)
  
  If strEstado <> "1" Then
    strEstado = "R"
     
    IdUsuario = ObtieneRutUsuarioLogueado()
    Vista = ObtenerVistaOrigen("CRS")
    Operacion = ObtenerOperacionBorrado("CRS", Mid$(DataINI, 60 + 26, 1))
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion + Mid$(DataINI, 1, 26), Mid$(DataINI, 27, 60))
  
  ElseIf VarSiebel = "NORMAL" Then
    
    strEstado = "A"
    '*** Arma Transaccion CRS, CRE (Credito Empresas PYME)
    IdUsuario = ObtieneRutUsuarioLogueado()
    Vista = ObtenerVistaDestino("CRS")
    Operacion = ObtenerOperacionSRM("CRS", Mid$(DataINI, 60 + 26, 1))
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion + Mid$(DataINI, 60 + 1, 26), Mid$(DataINI, 60 + 27, 465))
  
    Vista = ObtenerVistaOrigen("CRS")
    Operacion = ObtenerOperacionBorrado("CRS", Mid$(DataINI, 60 + 26, 1))
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion + Mid$(DataINI, 1, 26), Mid$(DataINI, 27, 60))
    
  Else '** SIEBEL
    strEstado = "A"
    If OperCRS = "O" Then
      OperCRS = "A"
    End If

    DataPLT = "b12256781234W00000012plt_emp01 exec spvsd_cre1217_ins "
  
    NomCliente = DatoClienteEncontrado(Mid$(DataINI, 51, 10))
    CodEjecutivoVenta = Mid$(Mid$(DataINI, 61), 249, 6)
    RutEjecutivo = TraeRutEjecutivo(CodEjecutivoVenta)
    
    Data1 = Mid$(DataINI, 61, 254)
    Data2 = Mid$(DataINI, 61 + 254, 238)
  
    Data3 = UCase(NomCliente) + RutEjecutivo
  
    For i = 1 To 7
      If ProductoAsociado_TraeRut(Encabezado + "SC" + CStr(i) + Llaveproducto, "1", "39", RutAsegurado) Then
         Data3 = Data3 + Left(RutAsegurado + Space(10), 10)
      Else
         Data3 = Data3 + Space(10)
      End If
    Next i
    
    Call ActualizarEnvioSRM(DataPLT + "'" + Data1 + "','", Data2 + "','" + Data3 + OperCRS + "'")
        
    '** borrado origen
    IdUsuario = ObtieneRutUsuarioLogueado
    Vista = ObtenerVistaOrigen("CRS")
    Operacion = ObtenerOperacionBorrado("CRS", Mid$(DataINI, 40, 1))
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion + Mid$(DataINI, 1, 26), Mid$(DataINI, 27, 60))
     
  End If
  
End If

BaseEnvioSrm.Close
BaseVisado.Close

P = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")

End

End Sub
