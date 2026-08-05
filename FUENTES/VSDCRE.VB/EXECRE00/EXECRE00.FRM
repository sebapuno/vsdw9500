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
Dim VarSiebel As String
Dim VarTO As String

If Not App.PrevInstance Then
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set BaseEnvioSrm = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
   
  RutCliente = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Correl = Val(SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01))
  VarSiebel = SacaValorINI("SIEBEL", "TIPOVSD", DIRECTORIO_01 + ARCHIVO_INI01)
  VarTO = SacaValorINI("SIEBEL", "CRETO", DIRECTORIO_01 + ARCHIVO_INI01)

  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
  PunteroVsdPrd.Seek "=", RutCliente, Producto, Correl
  DataINI = PunteroVsdPrd("Prd_Data")
  PunteroVsdPrd.Close
   
  strEstado = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)
  OperCRE = Mid$(DataINI, 26, 1)
  
  If strEstado <> 1 Then
  '*** SI RECHAZO TRANSACCIÓN *****
    
    strEstado = "R"
     
    IdUsuario = ObtieneRutUsuarioLogueado()
    Vista = ObtenerVistaOrigen("CRE")
    Operacion = ObtenerOperacionBorrado("CRE", Mid$(DataINI, 26, 1))
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion + Mid$(DataINI, 1, 26), Mid$(DataINI, 27, 465))
     
  ElseIf VarSiebel <> "NORMAL" And VarTO <> "NORMAL" Then
    '*** SI PC ES SIEBEL Y TO EXISTE EN INI REMOTO  *******
    
    '** Para que las transacciones no usen sat
    Set PunteroMsg = BaseEnvioSrm.OpenRecordset("Mensaje", dbOpenTable)
    PunteroMsg.MoveFirst
    PunteroMsg.Edit
    PunteroMsg!SAT = "0"
    PunteroMsg!TX_SAT = "XXXX"
    PunteroMsg.Update
    PunteroMsg.Close
    
    If OperCRE = "N" Then
      OperCRE = "A"
    End If
    
    NomCliente = UCase(DatoClienteEncontrado(Mid$(DataINI, 10, 10)))
    CodEjecutivoVenta = Mid$(DataINI, 249, 6)
    RutEjecutivo = TraeRutEjecutivo(CodEjecutivoVenta)
  
    DataPLT = "b12256781234W00000012plt_emp01 exec spvsd_cre1217_ins "
     
    Data1 = Mid$(DataINI, 1, 254)
    If Trim$(Mid$(DataINI, 255 + 221, 13)) <> "" Then
      Data2 = Mid$(DataINI, 255, 238)
    Else
      Data2 = Mid$(DataINI, 255, 221) + String(13, "0") + Mid$(DataINI, 255 + 234, 3)
    End If
            
    Data3 = NomCliente + RutEjecutivo + Space(70)
  
    Call ActualizarEnvioSRM(DataPLT + "'" + Data1 + "','", Data2 + "','" + Data3 + OperCRE + "'")
    
    strEstado = "A"
    '*** Arma Transaccion Modificar CRE (Credito Empresas PYME)
    IdUsuario = ObtieneRutUsuarioLogueado()
    Vista = ObtenerVistaDestino("CRE")
    Operacion = "W"
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion + Mid$(DataINI, 1, 26), Mid$(DataINI, 27, 464) + "L")
    
  Else
    '*** TRANSACCIÓN NORMAL
    
    strEstado = "A"
    '*** Arma Transaccion CRE (Credito Empresas PYME)
    IdUsuario = ObtieneRutUsuarioLogueado()
    Vista = ObtenerVistaDestino("CRE")
    Operacion = ObtenerOperacionSRM("CRE", Mid$(DataINI, 26, 1))
    Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion + Mid$(DataINI, 1, 26), Mid$(DataINI, 27, 465))
  
    '*** Arma Transaccion EXC (Exepciones al Visado)
    Call ExcepcionesAlVisado
     
  End If
End If

BaseEnvioSrm.Close
BaseVisado.Close

P = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")

End

End Sub
