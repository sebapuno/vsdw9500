VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   1320
   ClientTop       =   1425
   ClientWidth     =   3915
   Height          =   2220
   Left            =   1260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1815
   ScaleWidth      =   3915
   Top             =   1080
   Visible         =   0   'False
   Width           =   4035
   ' FI2026 - VBX.SpreadSheet migrado a MSFlexGrid (MSFLXGRD.OCX)
   Begin MSFlexGridLib.MSFlexGrid GrillaProducto 
      Height          =   1695
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3810
      Cols            =   2
      Rows            =   2
      ScrollBars      =   2  'flexScrollBarVertical
      SelectionMode   =   1  'flexSelectionByRow
   End
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Data As String * 700
Dim DataPatron As String
Dim BaseVisado As Database
Dim BaseSGT As Database
Public Function ColocarSeparadorDeMiles(ByVal Numero As String) As String
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
  
If PosicionComa > 0 Then
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
  ColocarSeparadorDeMiles = Cantidad + ";" + Decimal
Else
  ColocarSeparadorDeMiles = Numero
End If


End Function

Private Sub Form_Load()
Dim Respuesta As Integer
  If Not App.PrevInstance Then
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set BaseSGT = OpenDatabase(DIRECTORIO_02 + ARCHIVO_MDB02)
    GrillaProducto.Cols = 4
    Call LlenadoGrillaProducto
    Call LlenarArchivoTexto
    Respuesta = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
    BaseSGT.Close
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
Dim FiltroTabDesp As Recordset
Dim Producto As String
Dim Correlat As String
Dim LlaveProducto As String
Dim Num As String
Dim Moneda As String

RutCliente = SacaValorINI("Parametros", "PARAM1", DIRECTORIO_01 + ARCHIVO_INI01)
Producto = SacaValorINI("Parametros", "PARAM2", DIRECTORIO_01 + ARCHIVO_INI01)
Correlat = SacaValorINI("Parametros", "PARAM3", DIRECTORIO_01 + ARCHIVO_INI01)

Set Ap_VsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
' FI2026 - Ap_VsdPrd.Index = Ap_VsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
Ap_VsdPrd.Seek "=", RutCliente, Producto, Correlat

If Not Ap_VsdPrd.NoMatch Then
  Encabezado = Mid$(Ap_VsdPrd!Prd_Data, 1, 6)
  LlaveProducto = Mid$(Ap_VsdPrd!Prd_Data, 10, 30)

  Set FiltroTabDesp = BaseVisado.OpenRecordset("SELECT * FROM Tab_Desp WHERE Desp_Prod='MHB';")
  While Not FiltroTabDesp.EOF
    If FiltroTabDesp("Desp_Fil") > GrillaProducto.Rows Then
      GrillaProducto.Rows = FiltroTabDesp("Desp_Fil")
    End If
    GrillaProducto.Col = FiltroTabDesp("Desp_Col")
    GrillaProducto.Row = FiltroTabDesp("Desp_Fil")
    GrillaProducto.Text = FiltroTabDesp("Desp_Glosa")
    If FiltroTabDesp("Desp_Picture") <> "" Then
      GrillaProducto.Col = GrillaProducto.Col + 1
      GrillaProducto.Text = ObtenerDescripcionGlosa(Ap_VsdPrd!Prd_Data, FiltroTabDesp("Desp_Pos"), FiltroTabDesp("Desp_Picture"))
    End If
    FiltroTabDesp.MoveNext
  Wend
  FiltroTabDesp.Close

  '***COLOCA DATO DE MHX
  If ProductoAsociadoEncontrado(Encabezado + "MHX" + Mid$(LlaveProducto, 1, 22) + "001", "1", "34") Then
    FchMacAnterior = Mid$(DataPatron, 69, 8)
    If Trim$(FchMacAnterior) <> "" Then
      FchMacAnterior = Mid$(FchMacAnterior, 7, 2) + "/" + Mid$(FchMacAnterior, 5, 2) + "/" + Mid$(FchMacAnterior, 1, 4)
    End If
  
    GrillaProducto.Rows = GrillaProducto.Rows + 1
  
    GrillaProducto.Row = 3
    GrillaProducto.Col = 1
    GrillaProducto.Text = "Fecha MAC Anterior"
    GrillaProducto.Col = 2
    GrillaProducto.Text = FchMacAnterior
  
    GrillaProducto.Row = GrillaProducto.Rows
    GrillaProducto.Col = 1
    GrillaProducto.Text = "Detalles en OpciÃƒÂ³n Pauta"
  End If
  
End If

Ap_VsdPrd.Close

    
End Sub
Private Function ObtenerDescripcionGlosa(DatoEntrada As String, DondeCortar As String, Formato As String) As String
Dim Largo As Integer
Dim Posicion As Integer
Dim PosicionComa As Integer
Dim Seccion As String
Dim Signo As String * 1
Dim Numero As String
Dim TipoFormato As String
Dim IndiceTabGlosa As Recordset
Dim IndiceTblAec As Recordset

If DondeCortar <> "" Then
  PosicionComa = InStr(DondeCortar, ",")
  Posicion = Val(Mid(DondeCortar, 1, PosicionComa - 1))
  Largo = Val(Mid(DondeCortar, PosicionComa + 1))
  Seccion = Mid(DatoEntrada, Posicion, Largo)
Else
  Seccion = Texto
End If

Select Case Mid$(Formato, 1, 2)
  Case "TG"
    Set IndiceTabGlosa = BaseVisado.OpenRecordset("Tab_Glosa", dbOpenTable)
    ' FI2026 - IndiceTabGlosa.Index = IndiceTabGlosa.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    IndiceTabGlosa.Seek "=", CodProd, Mid$(Formato, 3, 3), Seccion
    If Not IndiceTabGlosa.NoMatch Then
      ObtenerDescripcionGlosa = IndiceTabGlosa!Glosa_Descrip
    End If
    IndiceTabGlosa.Close
    
  Case "AE"
    Set IndiceTblAec = BaseSGT.OpenRecordset("TblAec", dbOpenTable)
    ' FI2026 - IndiceTblAec.Index = IndiceTblAec.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    IndiceTblAec.Seek "=", Seccion
    If Not IndiceTblAec.NoMatch Then
      ObtenerDescripcionGlosa = StrConv(Trim$(IndiceTblAec!AEC_AECNOM), 3)
      ObtenerDescripcionGlosa = ReemplazaComas(ObtenerDescripcionGlosa)
    End If
    IndiceTblAec.Close
        
  Case "DC"
    Numero = Decodifica(Trim$(Seccion), Interpreta(Mid$(Formato, 3, 7), SIMBOLOPOSICION), _
                                         Interpreta(Mid$(Formato, 3, 7), SIMBOLOLARGO))
    Signo = Mid$(Numero, 1, 1)
    Numero = Mid$(Numero, 2, Len(Numero) - Interpreta(Mid$(Formato, 3, 7), SIMBOLOLARGO) - 1) + "," + _
                      Mid$(Numero, Len(Numero) - Interpreta(Mid$(Formato, 3, 7), SIMBOLOLARGO) + 1)
    While Mid$(Numero, 1, 1) = "0"
      Numero = Mid$(Numero, 2)
    Wend
    
    ObtenerDescripcionGlosa = Signo + ColocarSeparadorDeMiles(Numero)
      
  Case "VF"
      TipoFormato = Trim$(Mid(Formato, 3, Len(Formato)))
      Select Case TipoFormato
        Case "aaaammdd"
          ObtenerDescripcionGlosa = Mid$(Seccion, 7, 2) + "/" + Mid$(Seccion, 5, 2) + "/" + Mid$(Seccion, 1, 4)
        Case "ddmmssaa", "ddmmaaaa"
          ObtenerDescripcionGlosa = Mid$(Seccion, 1, 2) + "/" + Mid$(Seccion, 3, 2) + "/" + Mid$(Seccion, 5, 4)
      End Select
  Case Else
      Select Case Formato
         Case "RUT"
            ObtenerDescripcionGlosa = Format(Val(Mid$(Seccion, 1, Len(Seccion) - 1)), "###,###,###") + "-" + Right(Seccion, 1)
         Case Else
            ObtenerDescripcionGlosa = Seccion
      End Select
      
  End Select


End Function
Public Function Interpreta(Cadena As String, DatoBuscado As String) As Integer
Dim PosicionComa As Integer
  PosicionComa = InStr(Cadena, ",")
  If PosicionComa <> 0 Then
    If DatoBuscado = SIMBOLOPOSICION Then
      Interpreta = Val(Mid(Cadena, 1, PosicionComa - 1))
    Else
      Interpreta = Val(Mid(Cadena, PosicionComa + 1, Len(Cadena)))
    End If
  Else
    Interpreta = -1
  End If
End Function
Public Function Decodifica(Numero As String, Posicion As Integer, Largo As Integer) As String
Dim Signo As String * 1
Dim Numerito As String
Dim LargoTotal  As Integer
Dim PosicionInicio As Integer
  If Numero <> "" Then
    If Largo <= 0 Then
      PosicionInicio = 1
      LargoTotal = Posicion - 1
    Else
      PosicionInicio = Posicion + 1
      LargoTotal = Largo - 1
    End If
    If InStr("{ABCDEFGHI", Mid$(Numero, Len(Numero), 1)) <> 0 Then
      Signo = ""
    Else
      Signo = "-"
    End If
    Select Case Mid$(Numero, Len(Numero), 1)
      Case "{", "}"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "0"
      Case "A", "J"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "1"
      Case "B", "K"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "2"
      Case "C", "L"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "3"
      Case "D", "M"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "4"
      Case "E", "N"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "5"
      Case "F", "O"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "6"
      Case "G", "P"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "7"
      Case "H", "Q"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "8"
      Case "I", "R"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "9"
    End Select
    If Largo <= 0 Then
      Decodifica = Signo + Format$(Numerito, "0000000000000")
    Else
      Decodifica = Signo + Format$(Mid(Numero, 1, Posicion), "0000000000000") + Numerito
    End If
  Else
    Decodifica = " " + String(13, "0")
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
  For I = 1 To GrillaProducto.Rows
    Linea = ""
    GrillaProducto.Row = I
    For J = 1 To GrillaProducto.Cols
      GrillaProducto.Col = J
      Linea = Linea + Trim$(GrillaProducto.Text) + "Ã‚Â¬"
    Next J
    Do While Mid$(Linea, Len(Linea), 1) = "Ã‚Â¬"
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
    DataPatron = FiltroVsdPrd("Prd_Data")
    ProductoAsociadoEncontrado = True
  End If
  FiltroVsdPrd.Close
End Function


Public Function MedioPago(CodPago As String) As String

Select Case CodPago
  Case "CTD"
     MedioPago = "Cuenta Corriente"
  Case "CRE"
     MedioPago = "Cuenta Contable CRE"
  Case Else
     MedioPago = CodPago
End Select

End Function

Public Function TraeMoneda(CodMon As String) As String

Select Case CodMon
  Case "001"
     TraeMoneda = "$"
  Case "098"
     TraeMoneda = "UF"
  Case "011"
     TraeMoneda = "US$"
  Case Else
     TraeMoneda = "(" + CodMon + ")"
End Select

End Function

Public Function ReemplazaComas(DataNormal As String) As String

Dim DataB As String

DataB = DataNormal

While InStr(DataB, ",") > 0
  DataB = Mid$(DataB, 1, InStr(DataB, ",") - 1) + ";" + Mid$(DataB, InStr(DataB, ",") + 1, dbOpenSnapshot)
Wend

ReemplazaComas = DataB

End Function
