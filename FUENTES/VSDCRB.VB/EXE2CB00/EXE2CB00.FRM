VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1650
   ClientLeft      =   1665
   ClientTop       =   2460
   ClientWidth     =   3390
   Height          =   2055
   Left            =   1605
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   3390
   Top             =   2115
   Visible         =   0   'False
   Width           =   3510
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario As String * 8
Dim DataINI As String
Dim StrEstado As String * 1
Dim Vista As String * 4
Dim Operacion As String * 1



Sub ArmaAmbosSeguros(DataCR1 As String, DataCR5 As String)
Dim DataR   As String
Dim DataReg As String
Dim DataCR2 As String
Dim DataCR3 As String
Dim DataCR6 As String
Dim DataCR7 As String

      '**CR1 SEGURO DE VIDA
      DataR = LimpiaData(DataCR1)
      DataReg = Mid$(DataR, 61)
      Call ActualizarEnvioSRM("0" + IdUsuario, DataReg)
  
      'No lleva Transaccion STB (CR4)
      Call ActualizarEnvioSRM(DataSTB, "")
     
      '** CR3 RELACION CLIENTE-PRODUCTO SEGURO VIDA
      DataCR3 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR3" + Mid$(DataINI, 10, 30))
      Vista = ObtenerVistaDestino("2RV")
      Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR3, 40, 1))
      DataReg = Mid$(DataCR3, 61)
      Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)
     
      '** CR2 DIRECCION SEGURO VIDA
      DataCR2 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR2" + Mid$(DataINI, 10, 30))
      Vista = ObtenerVistaDestino("2DV")
      Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR2, 40, 1))
      DataReg = Mid$(DataCR2, 61)
      Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)
      
      '** CR5 SEGURO CESANTIA
      DataR = LimpiaData(DataCR5)
      DataReg = Mid$(DataR, 61)
      Call ActualizarEnvioSRM("0" + IdUsuario, DataReg)

      '** CR6 RELACION CLIENTE-PRODUCTO SEGURO CESANTIA
      DataCR6 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR6" + Mid$(DataINI, 10, 30))
      Vista = ObtenerVistaDestino("2RS")
      Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR6, 40, 1))
      DataReg = Mid$(DataCR6, 61)
      Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)
     
      '** CR7 DIRECCION SEGURO CESANTIA
      DataCR7 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR7" + Mid$(DataINI, 10, 30))
      Vista = ObtenerVistaDestino("2DC")
      Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR7, 40, 1))
      DataReg = Mid$(DataCR7, 61)
      Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)
      
End Sub

Sub ArmaNingunSeguro()

Call ActualizarEnvioSRM("NULL", "")
Call ActualizarEnvioSRM(DataSTB, "")
Call ActualizarEnvioSRM("NULL##", "NULL##NULL##NULL##NULL")

End Sub


Sub ArmaSoloSeguroCesantia(DataCR5 As String)
Dim DataR   As String
Dim DataReg As String
Dim DataCR6 As String
Dim DataCR7 As String

Call ActualizarEnvioSRM("NULL", "")       'No hay Seguro de Vida
Call ActualizarEnvioSRM(DataSTB, "")      'No hay Transaccion STB
Call ActualizarEnvioSRM("NULL##", "NULL") 'No Clt-Prd y Direccion de Seguro Vida
     
'** CR5 SEGURO CESANTIA
DataR = LimpiaData(DataCR5)
DataReg = Mid$(DataR, 61)
Call ActualizarEnvioSRM("0" + IdUsuario, DataReg)

'** CR6 RELACION CLIENTE-PRODUCTO SEGURO CESANTIA
DataCR6 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR6" + Mid$(DataINI, 10, 30))
Vista = ObtenerVistaDestino("2RS")
Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR6, 40, 1))
DataReg = Mid$(DataCR6, 61)
Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)
    
'** CR7 DIRECCION SEGURO CESANTIA
DataCR7 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR7" + Mid$(DataINI, 10, 30))
Vista = ObtenerVistaDestino("2DC")
Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR7, 40, 1))
DataReg = Mid$(DataCR7, 61)
Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)

End Sub

Sub ArmaSoloSeguroVida(DataCR1 As String)

Dim DataR   As String
Dim DataReg As String
Dim DataCR3 As String
Dim DataCR2 As String

  '**CR1 SEGURO DE VIDA
  DataR = LimpiaData(DataCR1)
  DataReg = Mid$(DataR, 61)
  Call ActualizarEnvioSRM("0" + IdUsuario, DataReg)
  
  'No lleva Transaccion STB (CR4)
  Call ActualizarEnvioSRM(DataSTB, "")
     
  '** CR3 RELACION CLIENTE-PRODUCTO SEGURO VIDA
  DataCR3 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR3" + Mid$(DataINI, 10, 30))
  Vista = ObtenerVistaDestino("2RV")
  Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR3, 40, 1))
  DataReg = Mid$(DataCR3, 61)
  Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)
     
  '** CR2 DIRECCION SEGURO VIDA
  DataCR2 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR2" + Mid$(DataINI, 10, 30))
  Vista = ObtenerVistaDestino("2DV")
  Operacion = ObtenerOperacionSRM("2CB", Mid$(DataCR2, 40, 1))
  DataReg = Mid$(DataCR2, 61)
  Call ActualizarEnvioSRM(IdUsuario + Vista + Operacion, DataReg)
     
  'No lleva transacciones de Seguro Cesantia
  Call ActualizarEnvioSRM("NULL##", "NULL##NULL")
  
End Sub

Private Function BorrarEnVistaOrigen(ByVal Data As String) As Boolean
Dim Id As String
Dim DataBorrado As String
  BorrarEnVistaOrigen = True
  
  Id = "b12256781234W0000001Nplt_per01 exec vsd_modifica_estado "
  DataBorrado = "'" + Mid$(Data, 1, 3) + "'," + "'" + Mid$(Data, 4, 3) + "'," _
                + "'" + Mid$(Data, 7, 3) + "'," + "'" + Mid$(Data, 10, 30) + "'," _
                + "'" + Mid$(Data, 40, 1) + "'," + "'" + Mid$(Data, 41, 2) + "'," _
                + "'" + Mid$(Data, 43, 8) + "'," + "'" + Mid$(Data, 51, 10) + "'," _
                + "'" + StrEstado + "'"
  Call ActualizarEnvioSRM(Id, DataBorrado)

End Function

Private Sub ActualizarEnvioSRM(ByVal LlaveEnvio As String, ByVal DataEnvio As String)
Dim BaseSRM As Database
Dim PunteroEnvioSRM As Recordset
  Set BaseSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroEnvioSRM = BaseSRM.OpenRecordset("Mensaje")
  PunteroEnvioSRM.MoveLast
  PunteroEnvioSRM.Edit
  If IsNull(PunteroEnvioSRM("Data")) Then
    PunteroEnvioSRM("Data") = LlaveEnvio + DataEnvio + "##"
  Else
    PunteroEnvioSRM("Data") = PunteroEnvioSRM("Data") + LlaveEnvio + DataEnvio + "##"
  End If
  PunteroEnvioSRM.Update
  PunteroEnvioSRM.Close
  BaseSRM.Close
  
End Sub

Public Function BorrarEnVistaOrigen_Directo(Data As String) As Boolean

Dim BaseSRM As Database
Dim PunteroEnvioSRM As Recordset
Dim Id As String
Dim DataBorrado As String

BorrarEnVistaOrigen_Directo = True
 
Id = "b12256781234W0000001Nplt_per01 exec vsd_modifica_estado_STB "
DataBorrado = "'" + Mid$(Data, 1, 3) + "'," + "'" + Mid$(Data, 4, 3) + "'," _
                + "'" + Mid$(Data, 7, 3) + "'," + "'" + Mid$(Data, 10, 30) + "'," _
                + "'" + Mid$(Data, 40, 1) + "'," + "'" + Mid$(Data, 41, 2) + "'," _
                + "'" + Mid$(Data, 43, 8) + "'," + "'" + Mid$(Data, 51, 10) + "'," _
                + "'E'"

Set BaseSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
Set PunteroEnvioSRM = BaseSRM.OpenRecordset("Mensaje")
PunteroEnvioSRM.MoveLast
PunteroEnvioSRM.Edit
If IsNull(PunteroEnvioSRM!Directo) Then
  PunteroEnvioSRM!Directo = Id + DataBorrado + "##"
Else
  PunteroEnvioSRM!Directo = PunteroEnvioSRM!Directo + Id + DataBorrado + "##"
End If
PunteroEnvioSRM.Update
PunteroEnvioSRM.Close
BaseSRM.Close

End Function

Private Function ObtieneRutUsuarioLogueado() As String
Dim Rut As String * 70
Dim Retorno As Integer
  Retorno = AisGetUsr(Rut)
  If Retorno = 1 Then
    ObtieneRutUsuarioLogueado = Format$(Mid(Rut, 1, 8), "00000000")
  Else
    ObtieneRutUsuarioLogueado = RUTDEFECTO
  End If
End Function
Private Function ObtenerOperacionBorradoSRM(ByVal producto As String, ByVal Actividad As String) As String
Dim PunteroTabOper As Recordset

  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
  PunteroTabOper.Seek "=", producto, Actividad
  ObtenerOperacionBorradoSRM = LTrim$(PunteroTabOper("Oper_Borrado"))
  PunteroTabOper.Close
  
End Function

Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function SacaValorINI(ByVal Grupo As String, ByVal Titulo As String, ByVal NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(800, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Private Sub Form_Load()

Dim Rut As String
Dim Correlativo As String
Dim MiSql As String
Dim VsdPrd As Recordset
Dim Valor As Boolean

Valor = False
If Not App.PrevInstance Then
  
  IdUsuario = ObtieneRutUsuarioLogueado()
 
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set BaseEnvioSrm = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  
  Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  
   
  Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
  
  MiSql = "Select * From VSD_PRD Where PRD_CODPROD='2CB' And PRD_RUTCLI='" + Rut + _
                                                   "' And PRD_CORREL=" + Correlativo
  Set VsdPrd = BaseVisado.OpenRecordset(MiSql, dbOpenDynaset)
  If VsdPrd.RecordCount > 0 Then
    DataINI = VsdPrd!Prd_Data
  End If
  VsdPrd.Close
 
  StrEstado = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)
  If StrEstado = "1" Then
    StrEstado = "A"
    Valor = Arma_Data_Aceptada
  Else
    StrEstado = "R"
    Valor = Arma_Data_Borrada
  End If
  
  If Valor Then
    If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
       MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                    "Error de Escritura"
    End If
  End If
  BaseEnvioSrm.Close
  BaseVisado.Close
End If

End 'FIN DE LA APLICACIÓN
 
End Sub



Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal Posicion As String, ByVal Largo As String) As Boolean
Dim FiltroVsdPrd  As Snapshot
  ProductoAsociadoEncontrado = False
  Set FiltroVsdPrd = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Prd WHERE " + _
                           "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + PatronBusqueda + "';")
  If FiltroVsdPrd.RecordCount > 0 Then
    DataPatron = FiltroVsdPrd("Prd_Data")
    ProductoAsociadoEncontrado = True
  End If
  FiltroVsdPrd.Close
End Function


Public Function ObtenerVistaDestino(CodPrd As String) As String

Dim PunteroTabPrd As Recordset

Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
PunteroTabPrd.Index = PunteroTabPrd.Indexes(0).Name
PunteroTabPrd.Seek "=", CodPrd
ObtenerVistaDestino = LTrim$(PunteroTabPrd!Prd_Vista_Dst)
PunteroTabPrd.Close

End Function

Public Function ObtenerOperacionSRM(producto As String, Actividad As String) As String
Dim PunteroTabOper As Recordset

Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
PunteroTabOper.Seek "=", producto, Actividad
ObtenerOperacionSRM = LTrim$(PunteroTabOper!Oper_srm)
PunteroTabOper.Close

End Function


Public Function EncuentraAsociado(producto As String) As String

Dim FiltroVsdPrd  As Snapshot

EncuentraAsociado = False
Set FiltroVsdPrd = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Prd WHERE " + _
                        "Mid$(Prd_Data, 1, 39)='" + producto + "';")
If FiltroVsdPrd.RecordCount > 0 Then
  'FiltroVsdPrd.MoveLast
  'FiltroVsdPrd.MoveFirst
  EncuentraAsociado = FiltroVsdPrd("Prd_Data")
End If
FiltroVsdPrd.Close

End Function

Public Function LimpiaData(datito As String) As String

Dim pos As Integer
Dim texto As String

texto = datito
While InStr(texto, "@") > 0
   pos = InStr(texto, "@")
   texto = Mid(texto, 1, pos - 1) + " " + Mid(texto, pos + 1)
Wend
LimpiaData = texto

End Function

Public Function Arma_Data_Aceptada() As Boolean
Dim DataCR1 As String
Dim DataCR5 As String
Dim ExisteCR1 As Boolean
Dim ExisteCR5 As Boolean

'Modificacion 08/2005 JSA
'Se modifico toda la logica ya que existira solo un archivo vsd_0006 para CRB y CRF

ExisteCR1 = False
ExisteCR5 = False
DataSTB = "NULL##NULL##NULL##NULL##NULL##NULL##NULL##NULL##NULL"

'** Arma primera transaccion CREDITO EN CUOTAS
Call ActualizarEnvioSRM("", Mid(DataINI, 61)) '(IdUsuario + "CRP0" + "A", Mid$(DataINI, 61))
  
If ProductoAsociadoEncontrado(Mid$(DataINI, 1, 6) + "CR1" + Mid$(DataINI, 10, 30), "1", "39") Then
     '** CR1  SEGURO DE VIDA
     DataCR1 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR1" + Mid$(DataINI, 10, 30))
     ExisteCR1 = True
End If
  
If ProductoAsociadoEncontrado(Mid$(DataINI, 1, 6) + "CR5" + Mid$(DataINI, 10, 30), "1", "39") Then
     '** CR5  SEGURO DE CESANTIA
     DataCR5 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR5" + Mid$(DataINI, 10, 30))
     ExisteCR5 = True
End If
     
'Podrian darse los siguientes casos
'Solo S. Vida
'Solo S. Cesantia
'Ambos
'Ninguno

If ExisteCR1 And Not ExisteCR5 Then
  Call ArmaSoloSeguroVida(DataCR1)
Else
  If Not ExisteCR1 And ExisteCR5 Then
     Call ArmaSoloSeguroCesantia(DataCR5)
  Else
    If ExisteCR1 And ExisteCR5 Then
      Call ArmaAmbosSeguros(DataCR1, DataCR5)
    Else
      If Not ExisteCR1 And Not ExisteCR5 Then
        Call ArmaNingunSeguro
      End If
    End If
  End If
End If

'** Eliminacion del CRB
BorrarEnVistaOrigen (DataINI)

'** Eliminacion del CRB directamente a la base
BorrarEnVistaOrigen_Directo (DataINI)

Arma_Data_Aceptada = True

End Function

Public Function Arma_Data_Borrada() As Boolean

Dim PunteroMsg As Recordset
Dim DataResp As String
Dim DataReg As String
Dim DataR As String
Dim DataCR1 As String
Dim DataCR2 As String
Dim DataCR3 As String
Dim DataCR5 As String
Dim DataCR6 As String
Dim DataCR7 As String

Arma_Data_Borrada = False

'** Eliminacion del CRB
If BorrarEnVistaOrigen(DataINI) Then
  Arma_Data_Borrada = True
End If

If ProductoAsociadoEncontrado(Mid$(DataINI, 1, 6) + "CR1" + Mid$(DataINI, 10, 30), "1", "39") Then
  '** Eliminacion del CR1
  DataCR1 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR1" + Mid$(DataINI, 10, 30))
  If BorrarEnVistaOrigen(DataCR1) Then
    '** Eliminacion del CR3
    DataCR3 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR3" + Mid$(DataINI, 10, 30))
    If BorrarEnVistaOrigen(DataCR3) Then
      '** Eliminacion del CR2
      DataCR2 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR2" + Mid$(DataINI, 10, 30))
      If BorrarEnVistaOrigen(DataCR2) Then
         '** Eliminacion del CR5
         DataCR5 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR5" + Mid$(DataINI, 10, 30))
         If BorrarEnVistaOrigen(DataCR5) Then
           '** Eliminacion del CR6
           DataCR6 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR6" + Mid$(DataINI, 10, 30))
           If BorrarEnVistaOrigen(DataCR6) Then
             '** Eliminacion del CR7
             DataCR7 = EncuentraAsociado(Mid(DataINI, 1, 6) + "CR7" + Mid$(DataINI, 10, 30))
             If BorrarEnVistaOrigen(DataCR7) Then
                Arma_Data_Borrada = True
             End If
           End If
         End If
         
      End If
    End If
  End If
End If
  
'** Para que las transacciones no usen sat
Set PunteroMsg = BaseEnvioSrm.OpenRecordset("Mensaje", dbOpenTable)
PunteroMsg.MoveFirst
PunteroMsg.Edit
PunteroMsg!SAT = "0"
PunteroMsg!TX_SAT = "XXXX"
PunteroMsg.Update
PunteroMsg.Close

End Function
