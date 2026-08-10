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
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Private Sub LlenadoGrillaProducto()
Dim Encabezado As String
Dim FiltroTabDesp As Recordset
Dim FechaApertura As String
Dim NumeroCuenta As String
Dim ExisteDireccion As String
Dim ExisteBipersonal As String
Dim ExisteReposicion As String
  Encabezado = Mid$(Data, 1, 6)
  FechaApertura = Mid$(Data, 103, 4)
  NumeroCuenta = Mid$(Data, 11, 11)
  ExisteDireccion = "No"
  ExisteBipersonal = "Sin Bipersonal"
  ExisteReposicion = "No"
  If ProductoAsociadoEncontrado(Encabezado + "CCR" + NumeroCuenta + "0", "1", "21") Then
    ExisteReposicion = "Si"
  End If
  Set FiltroTabDesp = BaseVisado.OpenRecordset("SELECT * FROM Tab_Desp WHERE Desp_Prod='1DD';")
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
                                                                                     ExisteDireccion, ExisteBipersonal, FechaApertura)
    End If
    FiltroTabDesp.MoveNext
  Wend
  FiltroTabDesp.Close
End Sub
Private Function ObtenerDescripcionGlosa(ByVal DondeCortar As String, ByVal Formato As String, ByVal GlosaReposicion As String, ByVal GlosaDireccion As String, ByVal GlosaBipersonal As String, ByVal GlosaFecha As String) As String
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
    Case "FEC"
      ObtenerDescripcionGlosa = Format$(Mid(GlosaFecha, 3, 2) + Mid(GlosaFecha, 1, 2), Formato)
    Case Else
      PosicionComa = InStr(DondeCortar, ",")
      Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
      Largo = Val(Mid(DondeCortar, PosicionComa + 1))
      ObtenerDescripcionGlosa = Format$(Mid(Data, Posicion, Largo), Trim(Formato))
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
                                                                                 "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + _
                                                                                 PatronBusqueda + "';", dbOpenSnapshot)
  If FiltroVsdPrd.RecordCount > 0 Then
    DataPatron = FiltroVsdPrd("Prd_Data")  ' FI2026 - se saca el dbOpenSnapshot que un --autofix pego al acceso al campo
    ProductoAsociadoEncontrado = True
  End If
  FiltroVsdPrd.Close
End Function


