VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   2805
   ClientTop       =   1440
   ClientWidth     =   3270
   Height          =   1800
   Left            =   2745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1395
   ScaleWidth      =   3270
   Top             =   1095
   Visible         =   0   'False
   Width           =   3390
   ' FI2026 - VBX.SpreadSheet migrado a MSFlexGrid (MSFLXGRD.OCX)
   Begin MSFlexGridLib.MSFlexGrid GrillaProducto 
      Height          =   1320
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3210
      Cols            =   1
      Rows            =   1
      ScrollBars      =   2  'flexScrollBarVertical
      SelectionMode   =   1  'flexSelectionByRow
   End
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Data As String
Dim DataPatron As String
Dim BaseVisado As Database
Public Function Busca_NomCom(CodCom As String) As String

Dim BaseSGT As Database
Dim p_Com As Recordset

If Trim$(CodCom) <> "" Then
  Set BaseSGT = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
  Set p_Com = BaseSGT.OpenRecordset("Select * from TBLCOM Where com_comcod=" + _
                                   Mid$(CodCom, 1, 3) + "And com_comdis=0", dbOpenDynaset)
  If p_Com.RecordCount > 0 Then
     Busca_NomCom = p_Com!Com_ComNom
  Else
     Busca_NomCom = ""
  End If
  p_Com.Close
  BaseSGT.Close
End If

End Function
Private Sub Form_Load()
Dim Respuesta As Integer
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Data = SacaValorINI("Parametros", "LlaveVisad", DIRECTORIO_01 + ARCHIVO_INI01)
    GrillaProducto.Cols = 5
    ' FI2026 - el Spread original no mostraba encabezados (DisplayColHeaders=0):
    ' MSFlexGrid siempre reserva la fila/columna 0, se ocultan con tamano 0
    GrillaProducto.RowHeight(0) = 0
    GrillaProducto.ColWidth(0) = 0
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
  lpReturnedString = String$(500, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Private Sub LlenadoGrillaProducto()
Dim RutCliente As String
Dim Encabezado As String
Dim NumeroCuenta As String
Dim FiltroTabDesp As Recordset
Dim ExisteDireccion As String
Dim ExisteBipersonal As String
Dim ExisteReposicion As String
Dim p_VsdPrd As Recordset
Dim LargoD As String

  RutCliente = Mid$(Data, 51, 9)
  Encabezado = Mid$(Data, 1, 6)
  NumeroCuenta = Mid$(Data, 10, 11)
  ExisteDireccion = "No"
  ExisteBipersonal = "Sin Bipersonal"
  ExisteReposicion = "No"
  If ProductoAsociadoEncontrado(Encabezado + "REC" + NumeroCuenta, "1", "20") Then
    ExisteReposicion = "Si"
  End If
  If ProductoAsociadoEncontrado(Encabezado + "DR3" + NumeroCuenta, "1", "33") Then
    DataPatron = Mid$(DataPatron, 51, 13)
    If Trim$(DataPatron) <> "" Then
      If ProductoAsociadoEncontrado(Encabezado + "DR2" + RutCliente + "CTD" + NumeroCuenta, "1", "33") Then
        ExisteDireccion = "Si"
      End If
    Else
      ExisteDireccion = "Si"
    End If
  End If
  If ProductoAsociadoEncontrado(Encabezado + "ABI" + "0" + NumeroCuenta, "1", "21") Then
    DataPatron = Mid$(DataPatron, 61, 10)
    ExisteBipersonal = Mid$(DataPatron, 1, 3) + "." + Mid$(DataPatron, 4, 3) + "." + _
                                    Mid$(DataPatron, 7, 3) + "-" + UCase$(Mid(DataPatron, 10, 1))
    While InStr("123456789", Mid$(ExisteBipersonal, 1, 1)) = 0
      ExisteBipersonal = Mid$(ExisteBipersonal, 2)
    Wend
  End If
  Set FiltroTabDesp = BaseVisado.OpenRecordset("SELECT * FROM Tab_Desp WHERE Desp_Prod='2RH';")
  While Not FiltroTabDesp.EOF
    If FiltroTabDesp("Desp_Fil") > GrillaProducto.Rows - 1 Then
      GrillaProducto.Rows = FiltroTabDesp("Desp_Fil") + 1
    End If
    GrillaProducto.Col = FiltroTabDesp("Desp_Col")
    GrillaProducto.Row = FiltroTabDesp("Desp_Fil")
    GrillaProducto.Text = FiltroTabDesp("Desp_Glosa")
    If FiltroTabDesp("Desp_Picture") <> "" Then
      GrillaProducto.Col = GrillaProducto.Col + 1
      GrillaProducto.Text = ObtenerDescripcionGlosa(FiltroTabDesp("Desp_Pos"), _
                                          FiltroTabDesp("Desp_Picture"), ExisteReposicion, _
                                          ExisteDireccion, ExisteBipersonal)
    End If
    FiltroTabDesp.MoveNext
  Wend
  FiltroTabDesp.Close
  
  '** Incorpora LEGAL
  GrillaProducto.Rows = GrillaProducto.Rows + 3
  GrillaProducto.Col = 1
  GrillaProducto.Row = GrillaProducto.Rows - 1
  GrillaProducto.Text = "DIRECCION LEGAL"
  GrillaProducto.Row = GrillaProducto.Rows - 1
  GrillaProducto.Text = "Dirección"
  LargoD = Str(Len(Encabezado + "LGL" + "CTD" + NumeroCuenta))
  Set p_VsdPrd = BaseVisado.OpenRecordset("Select * From VSD_PRD Where Prd_RutCli='" + Mid$(Data, 51, 10) + "' And " + _
                 "Mid$(Prd_Data, 1," + LargoD + ")='" + Encabezado + "LGLCTD" + NumeroCuenta + "'", dbOpenDynaset)
  
  If p_VsdPrd.RecordCount > 0 Then
    GrillaProducto.Col = 2
    GrillaProducto.Row = GrillaProducto.Rows - 1
    GrillaProducto.Text = Mid$(p_VsdPrd!Prd_Data, 73, 50)
    GrillaProducto.Col = 3
    GrillaProducto.Text = "Comuna"
    GrillaProducto.Col = 4
    GrillaProducto.Text = Busca_NomCom(Mid$(p_VsdPrd!Prd_Data, 123, 5))
      
    GrillaProducto.Rows = GrillaProducto.Rows + 1
    GrillaProducto.Row = GrillaProducto.Rows - 1
    GrillaProducto.Col = 1
    GrillaProducto.Text = "Teléfono"
    GrillaProducto.Col = 2
    GrillaProducto.Text = Format$(Mid$(p_VsdPrd!Prd_Data, 138, 12), "0")
  End If
  p_VsdPrd.Close
  '***
  
End Sub
Private Function ObtenerDescripcionGlosa(ByVal DondeCortar As String, ByVal Formato As String, ByVal GlosaReposicion As String, ByVal GlosaDireccion As String, ByVal GlosaBipersonal As String) As String
Dim Largo As Integer
Dim Posicion As Integer
Dim PosicionComa As Integer
Dim IndiceTabGlosa As Recordset

  Select Case DondeCortar
    Case "REP"
      ObtenerDescripcionGlosa = GlosaReposicion
    Case "DIR"
      ObtenerDescripcionGlosa = GlosaDireccion
    Case "BIP"
      ObtenerDescripcionGlosa = GlosaBipersonal
    Case Else
      If Mid$(Formato, 1, 2) = "VF" Then
        PosicionComa = InStr(DondeCortar, ",")
        Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
        Largo = Val(Mid(DondeCortar, PosicionComa + 1))
        Datp = Mid$(Data, Posicion, Largo)
        TipoFormato = Trim$(Mid(Formato, 3, Len(Formato)))
        Select Case TipoFormato
          Case "aaaammdd"
            ObtenerDescripcionGlosa = Mid$(Datp, 7, 2) + "/" + Mid$(Datp, 5, 2) + "/" + Mid$(Datp, 1, 4)
          Case "aammdd"
            ObtenerDescripcionGlosa = Mid$(Datp, 5, 2) + "/" + Mid$(Datp, 3, 2) + "/" + Mid$(Datp, 1, 2)
          Case "ddmmaa"
            ObtenerDescripcionGlosa = Mid$(Datp, 1, 2) + "/" + Mid$(Datp, 3, 2) + "/" + Mid$(Datp, 5, 2)
          Case "ddmmaaaa"
            ObtenerDescripcionGlosa = Mid$(Datp, 1, 2) + "/" + Mid$(Datp, 3, 2) + "/" + Mid$(Datp, 5, 4)
          Case "ddmmxxaa"
            ObtenerDescripcionGlosa = Mid$(Datp, 1, 2) + "/" + Mid$(Datp, 3, 2) + "/" + Mid$(Datp, 7, 2)
        End Select
      ElseIf Mid$(Formato, 1, 2) = "TG" Then
          PosicionComa = InStr(DondeCortar, ",")
          Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
          Largo = Val(Mid(DondeCortar, PosicionComa + 1))
          Datos = Mid$(Data, Posicion, Largo)
          Set IndiceTabGlosa = BaseVisado.OpenRecordset("Tab_Glosa", dbOpenTable)
          ' FI2026 - era: IndiceTabGlosa.Index = IndiceTabGlosa.Indexes(0).Name
          ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
          ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
          ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
          IndiceTabGlosa.Index = BaseVisado.TableDefs("Tab_Glosa").Indexes(0).Name
          IndiceTabGlosa.Seek "=", "2RH", Mid$(Formato, 3, 3), Datos
          If Not IndiceTabGlosa.NoMatch Then
             ObtenerDescripcionGlosa = IndiceTabGlosa("Glosa_Descrip")
          End If
          IndiceTabGlosa.Close
      Else
        PosicionComa = InStr(DondeCortar, ",")
        Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
        Largo = Val(Mid(DondeCortar, PosicionComa + 1))
        ObtenerDescripcionGlosa = Format$(Mid(Data, Posicion, Largo), Trim(Formato))
      End If
  End Select
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
  For I = 1 To GrillaProducto.Rows - 1
    Linea = ""
    GrillaProducto.Row = I
    For J = 1 To GrillaProducto.Cols - 1
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
Dim FiltroVsdPrd  As Recordset
  ProductoAsociadoEncontrado = False
  Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE " + _
                                               "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + PatronBusqueda + "';", dbOpenSnapshot)
  If FiltroVsdPrd.RecordCount > 0 Then
    DataPatron = FiltroVsdPrd("Prd_Data")  ' FI2026 - se saca el dbOpenSnapshot que un --autofix pego al acceso al campo
    ProductoAsociadoEncontrado = True
  End If
  FiltroVsdPrd.Close
End Function

