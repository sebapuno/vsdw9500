VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1395
   ClientLeft      =   2490
   ClientTop       =   2475
   ClientWidth     =   3270
   Height          =   1740
   Left            =   2460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1395
   ScaleWidth      =   3270
   Top             =   2160
   Visible         =   0   'False
   Width           =   3330
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
      InterfaceDesigner=   "ROTORLC2.frx":0000
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
Private Function FormateaRut(Rut As String) As String
  If Len(Rut) > 9 Then
    Rut = Mid$(Rut, 2)
  End If
  Rut = String(9 - Len(Rut), "0") + Rut
  FormateaRut = Mid$(Rut, 1, 2) + "." + Mid$(Rut, 3, 3) + "." + Mid$(Rut, 6, 3) + "-" + UCase$(Mid$(Rut, 9, 1))
End Function
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
Dim RutAval As String
Dim Encabezado As String
Dim FiltroTabDesp As Snapshot
Dim NumeroCuenta As String
Dim ExisteReposicion As String
  RutAval = "No Corresponde"
  ExisteReposicion = "No Corresponde"
  Set FiltroTabDesp = BaseVisado.CreateSnapshot("SELECT * FROM Tab_Desp WHERE Desp_Prod='LC2';")
  While Not FiltroTabDesp.EOF
    If FiltroTabDesp("Desp_Fil") > GrillaProducto.MaxRows Then
      GrillaProducto.MaxRows = FiltroTabDesp("Desp_Fil")
    End If
    GrillaProducto.Col = FiltroTabDesp("Desp_Col")
    GrillaProducto.Row = FiltroTabDesp("Desp_Fil")
    GrillaProducto.Text = FiltroTabDesp("Desp_Glosa")
    If FiltroTabDesp("Desp_Picture") <> "" Then
      GrillaProducto.Col = GrillaProducto.Col + 1
      GrillaProducto.Text = ObtenerDescripcionGlosa(FiltroTabDesp("Desp_Pos"), FiltroTabDesp("Desp_Picture"), _
                                                                                     ExisteReposicion, RutAval)
    End If
    FiltroTabDesp.MoveNext
  Wend
  FiltroTabDesp.Close
End Sub
Private Function ObtenerDescripcionGlosa(ByVal DondeCortar As String, ByVal Formato As String, ByVal GlosaReposicion As String, ByVal GlosaRutAval As String) As String
Dim Largo As Integer
Dim Datos As String
Dim Posicion As Integer
Dim PosicionComa As Integer
  Select Case DondeCortar
    Case "REP"
      ObtenerDescripcionGlosa = GlosaReposicion
    Case "AVA"
      ObtenerDescripcionGlosa = GlosaRutAval
    Case Else
      PosicionComa = InStr(DondeCortar, ",")
      Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
      Largo = Val(Mid(DondeCortar, PosicionComa + 1))
      Datos = Mid$(Data, Posicion, Largo)
      Select Case Mid$(Formato, 1, 2)
        Case "ND"
          Datos = Mid$(Datos, 1, Len(Datos) - Val(Mid(Formato, 3, 3))) + "," + _
                        Mid$(Datos, Len(Datos) - Val(Mid(Formato, 3, 3)) + 1)
          While Mid$(Datos, 1, 1) = "0"
            Datos = Mid$(Datos, 2)
          Wend
          Datos = ColocarSeparadorDeMiles(Datos)
          If Mid$(Formato, 3, 3) = "0" Then
            Datos = Mid$(Datos, 1, InStr(Datos, ",") - 1)
          End If
          ObtenerDescripcionGlosa = Datos
        Case "VF"
          TipoFormato = Trim$(Mid(Formato, 3, Len(Formato)))
          Select Case TipoFormato
            Case "aaaammdd"
              ObtenerDescripcionGlosa = Mid$(Datos, 7, 2) + "/" + Mid$(Datos, 5, 2) + "/" + Mid$(Datos, 1, 4)
            Case "aammdd"
              ObtenerDescripcionGlosa = Mid$(Datos, 5, 2) + "/" + Mid$(Datos, 3, 2) + "/" + Mid$(Datos, 1, 2)
            Case "ddmmaa"
              ObtenerDescripcionGlosa = Mid$(Datos, 1, 2) + "/" + Mid$(Datos, 3, 2) + "/" + Mid$(Datos, 5, 2)
            Case "ddmmssaa"
              ObtenerDescripcionGlosa = Mid$(Datos, 1, 2) + "/" + Mid$(Datos, 3, 2) + "/" + Mid$(Datos, 5, 4)
            Case "ddmmxxaa"
              ObtenerDescripcionGlosa = Mid$(Datos, 1, 2) + "/" + Mid$(Datos, 3, 2) + "/" + Mid$(Datos, 7, 2)
          End Select
        Case Else
          ObtenerDescripcionGlosa = Format$(Datos, Trim(Formato))
      End Select
  End Select
End Function
Private Function ColocarSeparadorDeMiles(ByVal Numero As String) As String
Dim Entero As String
Dim Decimal As String
Dim Cantidad As String
Dim PosicionComa As Integer
  Numero = Trim$(Numero)
  If Mid$(Numero, 1, 1) = "," Then
    Numero = "0" + Numero
  End If
  If Mid$(Numero, Len(Numero), 1) = "," Then
    Numero = Numero + "0"
  End If
  PosicionComa = InStr(Numero, ",")
  Entero = Mid$(Numero, 1, PosicionComa - 1)
  Decimal = Mid$(Numero, PosicionComa + 1)
  Cantidad = ""
  For I% = 1 To Len(Entero)
    Cantidad = Mid$(Numero, Len(Entero) + 1 - I%, 1) + Cantidad
    If I% Mod 3 = 0 Then
      Cantidad = "." + Cantidad
    End If
  Next I%
  If Mid$(Cantidad, 1, 1) = "." Then
    Cantidad = Mid$(Cantidad, 2)
  End If
  ColocarSeparadorDeMiles = Cantidad + "," + Decimal
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
                                                                                 "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + _
                                                                                 PatronBusqueda + "';")
  If FiltroVsdPrd.RecordCount > 0 Then
    DataPatron = FiltroVsdPrd("Prd_Data")
    ProductoAsociadoEncontrado = True
  End If
  FiltroVsdPrd.Close
End Function


