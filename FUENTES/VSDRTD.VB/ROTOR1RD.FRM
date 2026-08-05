VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   2490
   ClientTop       =   2475
   ClientWidth     =   3270
   Height          =   1800
   Left            =   2430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1395
   ScaleWidth      =   3270
   Top             =   2130
   Visible         =   0   'False
   Width           =   3390
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
      InterfaceDesigner=   "ROTOR1RD.frx":0000
      Left            =   45
      MaxCols         =   0
      MaxRows         =   0
      OperationMode   =   1  'ReadOnly
      RowHeaderDisplay=   0  'Blank
      ScrollBars      =   0  'None
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Data As String
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
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Private Sub LlenadoGrillaProducto()
Dim RutCliente As String
Dim Encabezado As String
Dim NumeroCuenta As String
Dim FiltroTabDesp As Snapshot
Dim ExisteDireccion As String
Dim ExisteBipersonal As String
Dim ExisteReposicion As String
  RutCliente = Mid$(Data, 51, 9)
  Encabezado = Mid$(Data, 1, 6)
  NumeroCuenta = Mid$(Data, 11, 11)
  ExisteDireccion = "No"
  ExisteBipersonal = "Sin Bipersonal"
  ExisteReposicion = "No"
  If ProductoAsociadoEncontrado(Encabezado + "1NS" + NumeroCuenta, "1", "20") Then
    ExisteReposicion = "Si"
  End If
  If ProductoAsociadoEncontrado(Encabezado + "DR3" + NumeroCuenta, "1", "33") Then
     ExisteDireccion = "Solo Producto"
     If ProductoAsociadoEncontrado(Encabezado + "DR2" + RutCliente + "CTD" + NumeroCuenta, "1", "33") Then
        ExisteDireccion = "Si"
     End If
  End If
  If ProductoAsociadoEncontrado(Encabezado + "1EG" + "0" + NumeroCuenta, "1", "21") Then
    DataPatron = Mid$(DataPatron, 61, 10)
    ExisteBipersonal = Mid$(DataPatron, 1, 3) + "." + Mid$(DataPatron, 4, 3) + "." + _
                                    Mid$(DataPatron, 7, 3) + "-" + UCase$(Mid(DataPatron, 10, 1))
    While InStr("123456789", Mid(ExisteBipersonal, 1, 1)) = 0
      ExisteBipersonal = Mid$(ExisteBipersonal, 2)
    Wend
  End If
  Set FiltroTabDesp = BaseVisado.CreateSnapshot("SELECT * FROM Tab_Desp WHERE Desp_Prod='3AR';")
  While Not FiltroTabDesp.EOF
    If FiltroTabDesp("Desp_Fil") > GrillaProducto.MaxRows Then
      GrillaProducto.MaxRows = FiltroTabDesp("Desp_Fil")
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
End Sub
Private Function ObtenerDescripcionGlosa(ByVal DondeCortar As String, ByVal Formato As String, ByVal GlosaReposicion As String, ByVal GlosaDireccion As String, ByVal GlosaBipersonal As String) As String
Dim Largo As Integer
Dim Posicion As Integer
Dim PosicionComa As Integer
  Select Case DondeCortar
    Case "REP"
      ObtenerDescripcionGlosa = GlosaReposicion
    Case "DIR"
      ObtenerDescripcionGlosa = GlosaDireccion
    Case "BIP"
      ObtenerDescripcionGlosa = GlosaBipersonal
    Case Else
      PosicionComa = InStr(DondeCortar, ",")
      Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
      Largo = Val(Mid(DondeCortar, PosicionComa + 1))
      If Mid$(Formato, 1, 2) = "VF" Then
          ObtenerDescripcionGlosa = Mid$(Data, Posicion + 6, 2) + "/" + Mid$(Data, Posicion + 4, 2) + "/" + Mid$(Data, Posicion, 4)
      Else
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

