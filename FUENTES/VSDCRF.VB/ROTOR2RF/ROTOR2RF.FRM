VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   1320
   ClientTop       =   1425
   ClientWidth     =   3315
   Height          =   1845
   Left            =   1260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1440
   ScaleWidth      =   3315
   Top             =   1080
   Visible         =   0   'False
   Width           =   3435
   Begin VBX.SpreadSheet GrillaProducto 
      BorderStyle     =   0  'None
      ColHeaderDisplay=   0  'Blank
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1320
      InterfaceDesigner=   "ROTOR2RF.frx":0000
      Left            =   45
      MaxCols         =   0
      MaxRows         =   0
      OperationMode   =   1  'ReadOnly
      RowHeaderDisplay=   0  'Blank
      ScrollBars      =   0  'None
      ShadowColor     =   &H00C8D0D4&
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Data As String * 1000
Dim DataPatron As String
Dim BaseVisado As Database
Private Sub Form_Load()
Dim Respuesta As Integer
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Data = SacaValorINI("Parametros", "LlaveVisad", DIRECTORIO_01 + ARCHIVO_INI01)
    GrillaProducto.MaxCols = 4
    Call LlenadoGrillaProducto
    Call LlenarArchivoTexto
    Respuesta = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
    BaseVisado.Close
  End If
  End
End Sub
Private Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(1000, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Private Sub LlenadoGrillaProducto()
Dim RutCliente As String
Dim Encabezado As String
Dim FiltroTabDesp As Snapshot
Dim RegistroAsociado As Snapshot
Dim Detalle1 As String * 1000
Dim Detalle2 As String * 1000
Dim Correlat As String

Dim FilaInicio As Integer

RutCliente = SacaValorINI("Parametros", "PARAM1", DIRECTORIO_01 + ARCHIVO_INI01)
Correlat = SacaValorINI("Parametros", "PARAM3", DIRECTORIO_01 + ARCHIVO_INI01)
Encabezado = Mid$(Data, 1, 6)
numero_cuenta = Mid$(Data, 10, 30)
Detalle1 = Mid$(Data, 61)

' ** COLOCA DATOS DE CRF  **********
Set FiltroTabDesp = BaseVisado.CreateSnapshot("SELECT * FROM Tab_Desp WHERE Desp_Prod='2RF';")
While Not FiltroTabDesp.EOF
  If FiltroTabDesp!Desp_Fil > GrillaProducto.MaxRows Then
    GrillaProducto.MaxRows = FiltroTabDesp!Desp_Fil
  End If
  GrillaProducto.Col = FiltroTabDesp!Desp_Col
  GrillaProducto.Row = FiltroTabDesp!Desp_Fil
  GrillaProducto.Text = FiltroTabDesp!Desp_Glosa
  If FiltroTabDesp!Desp_Picture <> "" Then
    GrillaProducto.Col = GrillaProducto.Col + 1
    GrillaProducto.Text = ObtenerDescripcionGlosa(Data, FiltroTabDesp!Desp_Pos, "", FiltroTabDesp!Desp_Picture)
  End If
  FiltroTabDesp.MoveNext
Wend
FiltroTabDesp.Close

FilaInicio = GrillaProducto.MaxRows + 2

'***COLOCA DATOS DE CR1 , SI EXISTE
If ProductoAsociadoEncontrado(Encabezado + "CR1" + numero_cuenta, "1", "39") Then
  Databuscada = Encabezado + "CR1" + numero_cuenta
  GrillaProducto.MaxRows = GrillaProducto.MaxRows + 4
  Set RegistroAsociado = BaseVisado.CreateSnapshot("SELECT * FROM vsd_prd WHERE prd_codprod='2SV' and prd_rutcli='" + RutCliente + "';")
  If RegistroAsociado.RecordCount > 0 Then
     RegistroAsociado.MoveLast
     RegistroAsociado.MoveFirst
     For I = 1 To RegistroAsociado.RecordCount
       If Databuscada = Mid$(RegistroAsociado!prd_data, 1, 39) Then
         Detalle1 = Mid$(RegistroAsociado!prd_data, 61)
         Exit For
       End If
       RegistroAsociado.MoveNext
     Next I
  End If
  RegistroAsociado.Close
  Call LlenaDatosSeguro(FilaInicio, Data, "VIDA", Detalle1)
  FilaInicio = FilaInicio + 4
End If

'***COLOCA DATOS DE CR5 , SI EXISTE
If ProductoAsociadoEncontrado(Encabezado + "CR5" + numero_cuenta, "1", "39") Then
  Databuscada = Encabezado + "CR5" + numero_cuenta
  GrillaProducto.MaxRows = GrillaProducto.MaxRows + 4
  Set RegistroAsociado = BaseVisado.CreateSnapshot("SELECT * FROM vsd_prd WHERE prd_codprod='2SC' and prd_rutcli='" + RutCliente + "';")
  If RegistroAsociado.RecordCount > 0 Then
     RegistroAsociado.MoveLast
     RegistroAsociado.MoveFirst
     For I = 1 To RegistroAsociado.RecordCount
       If Databuscada = Mid$(RegistroAsociado!prd_data, 1, 39) Then
         Detalle2 = Mid$(RegistroAsociado!prd_data, 61)
         Exit For
       End If
       RegistroAsociado.MoveNext
     Next I
  End If
  RegistroAsociado.Close
  Call LlenaDatosSeguro(FilaInicio, Data, "CESANTIA", Detalle2)
  
End If

   
End Sub

Private Sub LlenaDatosSeguro(FilaIni As Integer, Data As String, TipoSeg As String, Detalle As String)
Dim Fecha1 As String * 10
Dim Fecha2 As String * 10

GrillaProducto.Row = FilaIni
GrillaProducto.Col = 1
GrillaProducto.Text = "Tipo de Seguro"
    
GrillaProducto.Col = 2
GrillaProducto.Text = TipoSeg

GrillaProducto.Col = 3
GrillaProducto.Text = "Rut Asegurado"
 
GrillaProducto.Col = 4
GrillaProducto.Text = ObtenerDescripcionGlosa("", "", Mid$(Data, 10, 9), "RUT")
     
GrillaProducto.Row = FilaIni + 1
GrillaProducto.Col = 1
GrillaProducto.Text = "Prima UF"
     
GrillaProducto.Col = 2
GrillaProducto.Text = ObtenerDescripcionGlosa("", "", Mid$(Detalle, 167, 13), "ND4")
     
GrillaProducto.Col = 3
GrillaProducto.Text = "Monto Asegurado"
  
GrillaProducto.Col = 4
GrillaProducto.Text = ObtenerDescripcionGlosa("", "", Mid$(Detalle, 96, 13), "ND4")
  
GrillaProducto.Row = FilaIni + 2
GrillaProducto.Col = 1
GrillaProducto.Text = "Vigencia"
  
Fecha1 = Mid$(Detalle, 112, 3) + Mid$(Detalle, 109, 3) + Mid$(Detalle, 115, 4)
Fecha2 = Mid$(Detalle, 122, 3) + Mid$(Detalle, 119, 3) + Mid$(Detalle, 125, 4)
GrillaProducto.Col = 2
GrillaProducto.Text = Fecha1 + " al " + Fecha2

End Sub
Private Function ObtenerDescripcionGlosa(DatoEntrada As String, DondeCortar As String, Texto As String, Formato As String) As String
Dim Largo As Integer
Dim Posicion As Integer
Dim PosicionComa As Integer
Dim Seccion As String

If DondeCortar <> "" Then
  PosicionComa = InStr(DondeCortar, ",")
  Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
  Largo = Val(Mid(DondeCortar, PosicionComa + 1))
  Seccion = Mid(DatoEntrada, Posicion, Largo)
Else
  Seccion = Texto
End If

If Mid$(Formato, 1, 2) = "TG" Then
  Set IndiceTabGlosa = BaseVisado.OpenRecordset("Tab_Glosa", dbOpenTable)
  IndiceTabGlosa.Index = IndiceTabGlosa.Indexes(0).Name
  IndiceTabGlosa.Seek "=", "2RF", Mid$(Formato, 3, 3), Seccion
  If Not IndiceTabGlosa.NoMatch Then
    ObtenerDescripcionGlosa = IndiceTabGlosa!Glosa_Descrip
  End If
  IndiceTabGlosa.Close
ElseIf Mid$(Formato, 1, 2) = "ND" Then
  deci = Val(Mid$(Formato, 3, 1))
  numero = Mid$(Seccion, 1, Len(Seccion) - deci)
  numero = Format(Val(numero), "###,###,###,##0")
  ObtenerDescripcionGlosa = numero + ";" + Mid$(Seccion, Len(Seccion) - deci + 1)
  ' , InStr(Seccion, ",") - 1) + ";" + Mid$(Seccion, InStr(Seccion, ",") + 1)
Else
  Select Case Formato
     Case "SN"
        ObtenerDescripcionGlosa = IIf(Seccion = "S", "Si", "No")
     Case "NUM", "###"
        ObtenerDescripcionGlosa = Format(Seccion, "###,###,###,##0")
     Case "VFaaaammdd"
        ObtenerDescripcionGlosa = Mid$(Seccion, 7, 2) + "/" + Mid$(Seccion, 5, 2) + "/" + Mid$(Seccion, 1, 4)
     Case "RUT"
        ObtenerDescripcionGlosa = Format(Val(Mid$(Seccion, 1, 8)), "###,###,###") + "-" + Mid$(Seccion, 9, 1)
     Case Else
        ObtenerDescripcionGlosa = Seccion
  End Select
End If

End Function
Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  Else
    PoneValorINI = False
  End If
End Function
Private Sub LlenarArchivoTexto()
Dim I As Integer
Dim J As Integer
Dim Linea As String
  Open DIRECTORIO_01 + ARCHIVO_TXT01 For Output As #1
  For I = 1 To GrillaProducto.MaxRows
    Linea = ""
    GrillaProducto.Row = I
    For J = 1 To GrillaProducto.MaxCols
      GrillaProducto.Col = J
      Linea = Linea + Trim$(GrillaProducto.Text) + "¬"
    Next J
    Do While Mid$(Linea, Len(Linea), 1) = "¬"
      Linea = Mid$(Linea, 1, Len(Linea) - 1)
      If Len(Linea) = 0 Then
        Exit Do
      End If
    Loop
    Print #1, Linea
  Next I
  Close #1
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

