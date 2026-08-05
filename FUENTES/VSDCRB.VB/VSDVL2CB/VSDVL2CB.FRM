VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1755
   ClientTop       =   2250
   ClientWidth     =   4350
   Height          =   2220
   Left            =   1695
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4350
   Top             =   1905
   Visible         =   0   'False
   Width           =   4470
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim Respuesta As Integer
Dim BaseVisado As Database
Dim PunteroVsdPrd  As Recordset
Dim DataIni As String * 50
Dim ipos As Integer
Dim Correlativo As String * 1
Dim rCR1, rCR2, rCR3 As Boolean

'Modif 12/07/2005 JSA
Dim rCR5, rCR6, rCR7 As Boolean

Dim Existe   As String
Dim Valido   As String
Dim CtaCargo As String
Dim CtaAbono As String

rCR1 = False
rCR2 = False
rCR3 = False

rCR5 = False
rCR6 = False
rCR7 = False

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2CB' and prd_rutcli ='" + Rut + "' and prd_correl=" + Correlativo, dbOpenSnapshot)
    DataIni = Mid$(PunteroVsdPrd!prd_data, 1, 50)
    
    CtaAbono = Mid$(PunteroVsdPrd!prd_data, 282, 20)
    CtaCargo = Mid$(PunteroVsdPrd!prd_data, 305, 20)
    
    '** Busca si existe registro : SEGURO DE VIDA ********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR1" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2SV' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR1 = True
              Exit For ' Encuentra CR1 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If
    '** Busca si existe registro : DIRECCCION SEGURO VIDA ********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR2" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2DV' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR2 = True
              Exit For ' Encuentra CR2 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If
    '** Busca si existe registro : RELACION CLIENTE-PRODUCTO SEGURO VIDA ********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR3" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2RV' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR3 = True
              Exit For ' Encuentra CR3 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If
    
    'Modif 12/07/2005 JSA
    '** Busca si existe registro : SEGURO DE CESANTIA ********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR5" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2SC' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR1 = True
              Exit For ' Encuentra CR5 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If
    
    '** Busca si existe registro : DIRECCCION SEGURO CESANTIA ********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR7" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2DC' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR2 = True
              Exit For ' Encuentra CR7 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If
    
    '** Busca si existe registro : RELACION CLIENTE-PRODUCTO SEGURO CESANTIA ********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR6" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2RS' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR3 = True
              Exit For ' Encuentra CR6 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If
    
    'Modif 12/07/2005 JSA
    '*************************************************
    If (rCR1 = rCR2 And rCR2 = rCR3) And (rCR5 = rCR6 And rCR6 = rCR7) Then
       If ValidarCuentas(CtaCargo, CtaAbono) Then
          Respuesta = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
       End If
    Else
        If Not rCR1 Then
           MsgBox "Falta registro asociado CR1.", 16, "Error de Validación"
        ElseIf Not rCR2 Then
           MsgBox "Falta registro asociado CR2.", 16, "Error de Validación"
        ElseIf Not rCR3 Then
           MsgBox "Falta registro asociado CR3.", 16, "Error de Validación"
        ElseIf Not rCR5 Then
           MsgBox "Falta registro asociado CR5.", 16, "Error de Validación"
        ElseIf Not rCR6 Then
           MsgBox "Falta registro asociado CR6.", 16, "Error de Validación"
        ElseIf Not rCR7 Then
           MsgBox "Falta registro asociado CR7.", 16, "Error de Validación"
        End If
    End If
    '**************************************************
    
    PunteroVsdPrd.Close
    BaseVisado.Close
  End If
  End
  
End Sub

Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  Else
    PoneValorINI = False
  End If
End Function



Public Function ValidarCuentas(cCargo As String, cAbono As String) As Boolean
Dim Cuenta As String

ValidarCuentas = True

If Mid$(cCargo, 1, 3) = "CTD" Then
   Cuenta = Mid$(cCargo, 10)
   Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + Cuenta, 1)
   Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
   Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
   If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
     ValidarCuentas = True
   Else
     MsgBox "La Cuenta Cargo " + Cuenta + " No Existe.", 16, "Error de Validación"
     ValidarCuentas = False
   End If
End If
If Mid$(cAbono, 1, 3) = "CTD" And ValidarCuentas = True Then
   Cuenta = Mid$(cAbono, 10)
   Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + Cuenta, 1)
   Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
   Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
   If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
     ValidarCuentas = True
   Else
     MsgBox "La Cuenta Abono " + Cuenta + " No Existe.", 16, "Error de Validación"
     ValidarCuentas = False
   End If
End If


End Function
