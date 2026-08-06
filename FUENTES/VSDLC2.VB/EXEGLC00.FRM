VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1800
   ClientTop       =   2265
   ClientWidth     =   3480
   Height          =   2160
   Left            =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3480
   Top             =   1950
   Visible         =   0   'False
   Width           =   3540
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim IdUsuario As String * 8
Dim CadenaDatos As String


Private Sub Form_Load()
Dim Ret As Integer
Dim Data As String
Dim Llave As String
Dim GetAis As String * 70
Dim Actividad As String * 1
Dim Correlativo As String
Dim BaseVisado As Database
Dim PunteroVsdPrd As Recordset
Dim PunteroTabPrd As Recordset
Dim FechaHoy As String
Dim Tx_SAT As String * 4

  If Not App.PrevInstance Then
    ' FI2026 SHG - AisGetUsr (WAIS.DLL 16-bit) deshabilitado
    'Ret = AisGetUsr(GetAis)
    IdUsuario = Format(Mid(GetAis, 1, 8), "00000000")
    SalirAplicacion = False
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    PunteroTabPrd.Index = "Tab_Prd"
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    Cod_Rechazo = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)
    PunteroTabPrd.Seek "=", Producto
    PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
      FechaHoy = "'" + Format$(Date, "mm/dd/yyyy") + " " + Format$(Time, "hh:mm") + "'"
      If Cod_Rechazo = "0" Then
         Cod_Rechazo = "908"
         Cod_Visado = "F"
      Else
         Cod_Rechazo = "NULL"
         Cod_Visado = "E"
      End If
      PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
      Llave = "b12256781234W00000011orabch    "
      Data = "exec tcg_proc_gest_visado " + _
             Mid$(PunteroVsdPrd("Prd_Data"), 51, 9) + "," + Mid$(PunteroVsdPrd("Prd_Data"), 10, 9) + _
             "," + Cod_Rechazo + ",'" + Cod_Visado + "'," + FechaHoy
      Call ActualizarEnvioSRM(Llave, Data)
      If Mid$(PunteroVsdPrd("Prd_Data"), 10, 2) = "56" Then
         Llave = "TMB"
      Else
         Llave = "TAR"
      End If
      If Cod_Rechazo = "0" Then
         Data = "ELIMINA DE VISADO LA ACTIVACION TARJETA "   ' GLOSA FRACASO
      Else
         Data = "ACTIVACION (VISADO) " + Space(20)           ' GLOSA EXITO
      End If
      Data = Data + Mid$(PunteroVsdPrd("Prd_Data"), 51, 10)  ' RUT CLI
      Data = Data + Mid$(PunteroVsdPrd("Prd_Data"), 61, 50)  ' NOMBRE CLI
      Data = Data + Mid$(PunteroVsdPrd("Prd_Data"), 1, 3)    ' COD OFICINA
      Data = Data + GlosaOficina(Mid$(PunteroVsdPrd("Prd_Data"), 1, 3)) ' NOMBRE OFICINA
      Data = Data + "000"                                    ' COD EJEC
      Data = Data + "POSVENTA (MESON)" + Space(14)           ' NOM EJEC
      Data = Data + Mid$(PunteroVsdPrd("Prd_Data"), 10, 9)   ' LLAVE TX
      Data = Data + Space(200)                               ' INFOR.ADIC
      Call ActualizarEnvioSRM(Llave, Data)
    
    PunteroTabPrd.Close
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
End Sub


Private Sub ActualizarEnvioSRM(Llave As String, Data As String)
Dim BaseSRM As Database
Dim PunteroEnvioSRM As Recordset
  Set BaseSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroEnvioSRM = BaseSRM.OpenRecordset("Mensaje")
  If PunteroEnvioSRM.RecordCount > 0 Then
    PunteroEnvioSRM.MoveLast
    PunteroEnvioSRM.Edit
    PunteroEnvioSRM("Data") = PunteroEnvioSRM("Data") + Llave + Data + "##"
  Else
    PunteroEnvioSRM.AddNew
    PunteroEnvioSRM("Data") = Llave + Data + "##"
  End If
  PunteroEnvioSRM.Update
  PunteroEnvioSRM.Close
  BaseSRM.Close
End Sub

