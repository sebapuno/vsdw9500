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
Dim rCR1, rCR2, rCR3, rCR4 As Boolean
Dim rCR5, rCR6, rCR7  As Boolean
Dim Existe   As String
Dim Valido   As String
Dim Cuenta As String

rCR1 = False
rCR2 = False
rCR3 = False
rCR4 = False

rCR5 = False
rCR6 = False
rCR7 = False

  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
    Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2RF' and prd_rutcli ='" + Rut + "' and prd_correl=" + Correlativo, dbOpenSnapshot)
    DataIni = Mid$(PunteroVsdPrd!prd_data, 1, 50)
    
    '** Busca si existe registro : SEGURO VIDA********
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
    '** Busca si existe registro : DIRECCCION SEGURO VIDA********
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
    '** Busca si existe registro : RELACION CLIENTE-PRODUCTO SEGURO VIDA********
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
    '** Busca si existe registro : DESEMBOLSO EN LINEA ********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR4" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2TB' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              Cuenta = Mid$(PunteroVsdPrd!prd_data, 85, 11)
              rCR4 = True
              Exit For ' Encuentra CR4 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If

    'Modif JSA 18/07/2005
    '** Busca si existe registro : SEGURO CESANTIA********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR5" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2SC' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR5 = True
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
              rCR7 = True
              Exit For ' Encuentra CR7 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If
    
     '** Busca si existe registro : RELACION CLIENTE-PRODUCTO SEGURO CESANTIA********
    Data_Buscada = Mid$(DataIni, 1, 6) + "CR6" + Mid$(DataIni, 10, 30)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("select * from Vsd_Prd where prd_codprod = '2RS' and prd_rutcli ='" + Rut + "'", dbOpenSnapshot)
    If PunteroVsdPrd.RecordCount > 0 Then
        PunteroVsdPrd.MoveLast
        PunteroVsdPrd.MoveFirst
        For ipos = 1 To PunteroVsdPrd.RecordCount
           If Mid$(PunteroVsdPrd!prd_data, 1, 39) = Data_Buscada Then
              rCR6 = True
              Exit For ' Encuentra CR6 Correpondiente
           Else
              PunteroVsdPrd.MoveNext
           End If
        Next ipos
    End If


    '*************************************************
    If Not rCR4 Then ' obligatorio
       MsgBox "Falta registro asociado CR4.", 16, "Error de Validación"
    ElseIf (rCR1 = rCR2 And rCR2 = rCR3) And (rCR5 = rCR6 And rCR6 = rCR7) Then
       Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_PROFI + " XXX" + Cuenta, 1)
       Existe = SacaValorINI("Datos", "Flag_Existencia", DIRECTORIO_03 + ARCHIVO_INI03)
       Valido = SacaValorINI("Datos", "Flag_Validez", DIRECTORIO_03 + ARCHIVO_INI03)
       If UCase(Existe) = "TRUE" And UCase(Valido) = "TRUE" Then
          Respuesta = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
       Else
          MsgBox "La Cuenta Corriente " + Cuenta + " asociada al Crédito no existe.", 16, "Error de Validación"
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


