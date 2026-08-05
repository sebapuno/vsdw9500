VERSION 4.00
Begin VB.Form Frm_Detalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visado - MAC Histórico Empresas B"
   ClientHeight    =   6300
   ClientLeft      =   1155
   ClientTop       =   1500
   ClientWidth     =   11130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   6705
   Left            =   1095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Top             =   1155
   Width           =   11250
   Begin Threed.SSFrame SSFrame1 
      Height          =   4830
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   10965
      _version        =   65536
      _extentx        =   19341
      _extenty        =   8520
      _stockprops     =   14
      caption         =   "Detalle de los Créditos"
      forecolor       =   8388608
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      font3d          =   1
      shadowstyle     =   1
      Begin VBX.SpreadSheet Grid 
         DisplayRowHeaders=   0   'False
         Enabled         =   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   2505
         InterfaceDesigner=   "PTAMHB00.frx":0000
         Left            =   150
         MaxCols         =   8
         MaxRows         =   9
         OperationMode   =   1  'ReadOnly
         RetainSelBlock  =   0   'False
         ScrollBars      =   2  'Vertical
         ShadowColor     =   &H00C8D0D4&
         TabIndex        =   4
         Top             =   675
         Width           =   10600
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Index           =   0
         Left            =   3975
         TabIndex        =   6
         Top             =   210
         Width           =   1905
         _version        =   65536
         _extentx        =   3360
         _extenty        =   820
         _stockprops     =   15
         caption         =   "    Ap. Sugerida Según      Segmento y EEFF"
         forecolor       =   16777215
         backcolor       =   8388608
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Index           =   1
         Left            =   6990
         TabIndex        =   7
         Top             =   210
         Width           =   1155
         _version        =   65536
         _extentx        =   2037
         _extenty        =   820
         _stockprops     =   15
         caption         =   "Aprobados Vigentes"
         forecolor       =   16777215
         backcolor       =   8388608
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Index           =   2
         Left            =   8145
         TabIndex        =   8
         Top             =   210
         Width           =   1155
         _version        =   65536
         _extentx        =   2037
         _extenty        =   820
         _stockprops     =   15
         caption         =   "    Deuda      Actual"
         forecolor       =   16777215
         backcolor       =   8388608
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   465
         Index           =   3
         Left            =   9300
         TabIndex        =   9
         Top             =   210
         Width           =   1155
         _version        =   65536
         _extentx        =   2037
         _extenty        =   820
         _stockprops     =   15
         caption         =   "Sometido a Aprobación"
         forecolor       =   16777215
         backcolor       =   8388608
      End
      Begin Threed.SSPanel pnl_Dolar 
         Height          =   240
         Left            =   1500
         TabIndex        =   12
         Top             =   3825
         Width           =   1440
         _version        =   65536
         _extentx        =   2540
         _extenty        =   423
         _stockprops     =   15
         caption         =   "pnl_Dolar"
         forecolor       =   8388608
         backcolor       =   13160660
      End
      Begin Threed.SSPanel pnl_UF 
         Height          =   240
         Left            =   3825
         TabIndex        =   13
         Top             =   3825
         Width           =   1440
         _version        =   65536
         _extentx        =   2540
         _extenty        =   423
         _stockprops     =   15
         caption         =   "pnl_UF"
         forecolor       =   8388608
         backcolor       =   13160660
      End
      Begin VB.Label Label1 
         Caption         =   "U.F."
         Height          =   270
         Index           =   1
         Left            =   3375
         TabIndex        =   15
         Top             =   3840
         Width           =   390
      End
      Begin VB.Label Label1 
         Caption         =   "US$ Observado"
         Height          =   315
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   3840
         Width           =   1215
      End
      Begin MSGrid.Grid Grid2 
         Height          =   1455
         Left            =   5895
         TabIndex        =   11
         Top             =   3225
         Width           =   4610
         _version        =   65536
         _extentx        =   8123
         _extenty        =   2566
         _stockprops     =   77
         backcolor       =   16777215
         rows            =   6
         cols            =   4
         fixedrows       =   0
         scrollbars      =   0
         highlight       =   0   'False
         mouseicon       =   "PTAMHB00.frx":0573
      End
      Begin MSGrid.Grid Grid1 
         Height          =   250
         Left            =   750
         TabIndex        =   10
         Top             =   3225
         Width           =   5145
         _version        =   65536
         _extentx        =   9075
         _extenty        =   450
         _stockprops     =   77
         backcolor       =   16777215
         rows            =   1
         cols            =   3
         fixedrows       =   0
         scrollbars      =   0
         highlight       =   0   'False
         mouseicon       =   "PTAMHB00.frx":058F
      End
   End
   Begin Threed.SSFrame Frame3D1 
      Height          =   1320
      Left            =   75
      TabIndex        =   1
      Top             =   4875
      Width           =   8220
      _version        =   65536
      _extentx        =   14499
      _extenty        =   2328
      _stockprops     =   14
      caption         =   " Comentarios "
      forecolor       =   8388608
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      font3d          =   1
      Begin VB.Label lb_Comentarios 
         BackStyle       =   0  'Transparent
         Caption         =   "lb_Comentarios"
         Height          =   795
         Left            =   225
         TabIndex        =   5
         Top             =   375
         Width           =   7755
      End
   End
   Begin Threed.SSCommand Aceptar 
      Height          =   615
      Left            =   8625
      TabIndex        =   3
      Top             =   5025
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Aceptar"
      forecolor       =   8421376
      font3d          =   3
      picture         =   "PTAMHB00.frx":05AB
   End
   Begin Threed.SSCommand Cancelar 
      Height          =   615
      Left            =   9750
      TabIndex        =   2
      Top             =   5025
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      font3d          =   3
      picture         =   "PTAMHB00.frx":0A35
   End
End
Attribute VB_Name = "Frm_Detalle"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim StrUF As String
Dim StrDolar As String




Private Sub Aceptar_Click()

p = PoneValorINI("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")

End

End Sub


Public Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Sub Cancelar_Click()

End

End Sub

Private Sub Form_Load()

Dim Rut As String
Dim Producto As String
Dim Correlativo As String
Dim Comentarios As String

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

Grid.UserResizeCol = 2
Grid.RowHeight(0) = 15
  
Prepara_Grids

Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)

Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)

Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)

If Not PunteroVsdPrd.NoMatch Then
  lb_Comentarios.Caption = Trim$(Mid$(PunteroVsdPrd!Prd_Data, 61, 250))
  Call MuestraDetalle(PunteroVsdPrd!Prd_Data)
  Call CalculaTotales
End If

PunteroVsdPrd.Close
BaseVisado.Close

Success = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZENOMOVE)


End Sub





Public Sub Prepara_Grids()

'** GRID 1 **
'Me.Visible = True
Grid1.ColWidth(0) = 3200
Grid1.ColWidth(1) = 790
Grid1.ColWidth(2) = 1090

Grid1.Col = 0
Grid1.Text = "Garantía y Monto Total Sugerido"
Grid1.FixedAlignment(0) = 2 ' centro
Grid1.ColAlignment(1) = 2 ' centro
Grid1.ColAlignment(2) = 1 ' Derecha

Grid1.Enabled = False

'** GRID 2 **
Grid2.ColWidth(0) = 1110 '1150 ' 960
Grid2.ColWidth(1) = 1140
Grid2.ColWidth(2) = 1140
Grid2.ColWidth(3) = 1135

Grid2.Col = 0
Grid2.Row = 0
Grid2.Text = "Total $"
Grid2.Row = 1
Grid2.Text = "Total US$"
Grid2.Row = 2
Grid2.Text = "Total UF"
Grid2.Row = 3
Grid2.Text = "R.Directo($)"
Grid2.Row = 4
Grid2.Text = "R.Vinculado($)"
Grid2.Row = 5
Grid2.Text = "R.TOTAL($)"

Grid2.ColAlignment(1) = 1
Grid2.ColAlignment(2) = 1
Grid2.ColAlignment(3) = 1

Grid2.Enabled = False

End Sub

Public Sub MuestraDetalle(DataProducto As String)

Dim Encabezado As String
Dim LlaveProducto As String
Dim I As Integer
Dim Cantidad As Integer
Dim Moneda As String

Encabezado = Mid$(DataProducto, 1, 6)
LlaveProducto = Mid$(DataProducto, 10, 30)

'***COLOCA DATOS DE MHX

Cantidad = Val(Mid$(DataProducto, 61 + 250, 3))
   
For I = 1 To Cantidad
  DataPatron = ""
      
  If Not ProductoAsociadoEncontrado(Encabezado + "MHX" + Mid$(LlaveProducto, 1, 22) + Format$(I, "000"), "1") Then
    Exit For
  End If
  
  If I = 1 Then
    StrUF = Mid$(DataPatron, 239, 12) + "," + Mid$(DataPatron, 251, 2)
    num = Format$(Val(Mid$(DataPatron, 239, 12)), "0")
    num = num + "," + Mid$(DataPatron, 251, 2) ' 4)
    pnl_UF.Caption = "$ " + ColocarSeparadorDeMiles(num)

    StrDolar = Mid$(DataPatron, 255, 12) + "," + Mid$(DataPatron, 267, 2)
    num = Format$(Val(Mid$(DataPatron, 255, 12)), "0")
    num = num + "," + Mid$(DataPatron, 267, 2) ' 4)
    pnl_Dolar.Caption = "$ " + ColocarSeparadorDeMiles(num)
    
    Grid2.Row = 4
    '** Riesgo_AprobadoVigente
    Grid2.Col = 1
    num = Format$(Val(Mid$(DataPatron, 271, 19)), "0")
    'num = num + "," + Mid$(DataPatron, 271, 2) ' 4)
    Grid2.Text = ColocarSeparadorDeMiles(num)
    
    '** Riesgo_SometidoAprobacion
    Grid2.Col = 2
    num = Format$(Val(Mid$(DataPatron, 294, 19)), "0")
    'num = num + "," + Mid$(DataPatron, 294, 2) ' 4)
    Grid2.Text = ColocarSeparadorDeMiles(num)
    
    '** Riesgo_DeudaActual
    Grid2.Col = 3
    num = Format$(Val(Mid$(DataPatron, 317, 19)), "0")
    'num = num + "," + Mid$(DataPatron, 317, 2) ' 4)
    Grid2.Text = ColocarSeparadorDeMiles(num)
    
  End If
  
  Grid.Row = I
  Grid.Col = 1
  Grid.Text = Mid$(DataPatron, 79, 3)
  Grid.TypeHAlign = 2 ' center
   
  Grid.Col = 2
  Grid.Text = StrConv(Trim$(Mid$(DataPatron, 82, 60)), 3)
      
  Grid.Col = 3
  Grid.Text = Trim$(Mid$(DataPatron, 142, 2))
  Grid.TypeHAlign = 2 ' center
 
  Grid.Col = 4
  num = Format$(Val(Mid$(DataPatron, 144, 19)), "0")
  'num = num + "," + Mid$(DataPatron, 163, 2) ' 4)
  Grid.Text = ColocarSeparadorDeMiles(num)
  Grid.TypeHAlign = 1 ' derecha
    
  Grid.Col = 5
  Moneda = TraeMoneda(Mid$(DataPatron, 167, 3))
  Grid.Text = Moneda
  Grid.TypeHAlign = 2 ' centro
  
  Grid.Col = 6
  num = Format$(Val(Mid$(DataPatron, 170, 19)), "0")
  If Moneda <> "" And Moneda <> "$" Then
    num = num + "," + Mid$(DataPatron, 189, 2) ' 4)
    Grid.Text = ColocarSeparadorDeMiles(num)
  ElseIf Moneda <> "" And Moneda = "$" Then
    Grid.Text = ColocarSeparadorDeMiles(num)
  Else
    Grid.Text = ""
  End If
  Grid.TypeHAlign = 1 ' derecha
  
  Grid.Col = 7
  num = Format$(Val(Mid$(DataPatron, 193, 19)), "0")
  If Moneda <> "" And Moneda <> "$" Then
    num = num + "," + Mid$(DataPatron, 212, 2) ' 4)
    Grid.Text = ColocarSeparadorDeMiles(num)
  ElseIf Moneda <> "" And Moneda = "$" Then
    Grid.Text = ColocarSeparadorDeMiles(num)
  Else
    Grid.Text = ""
  End If
  Grid.TypeHAlign = 1 ' derecha
  
  Grid.Col = 8
  num = Format$(Val(Mid$(DataPatron, 216, 19)), "0")
  If Moneda <> "" And Moneda <> "$" Then
    num = num + "," + Mid$(DataPatron, 235, 2) ' 4)
    Grid.Text = ColocarSeparadorDeMiles(num)
  ElseIf Moneda <> "" And Moneda = "$" Then
    Grid.Text = ColocarSeparadorDeMiles(num)
  Else
    Grid.Text = ""
  End If
  Grid.TypeHAlign = 1 ' derecha
    
Next I


End Sub
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
  ColocarSeparadorDeMiles = Cantidad + "," + Decimal
Else
  ColocarSeparadorDeMiles = Format$(Val(Numero), "###,###,###,###,##0")
End If


End Function


Public Function TraeMoneda(CodMon As String) As String

Select Case CodMon
  Case "000"
     TraeMoneda = ""
  Case "001"
     TraeMoneda = "$"
  Case "098"
     TraeMoneda = "UF"
  Case "002"
     TraeMoneda = "US$"
  Case Else
     TraeMoneda = "(" + CodMon + ")"
End Select

End Function
Private Function ProductoAsociadoEncontrado(ByVal PatronBusqueda As String, ByVal Posicion As String) As Boolean
Dim Largo As String
Dim FiltroVsdPrd  As Snapshot

  Largo = LTrim$(Str(Len(PatronBusqueda)))
  ProductoAsociadoEncontrado = False
  Set FiltroVsdPrd = BaseVisado.CreateSnapshot("SELECT Prd_Data FROM Vsd_Prd WHERE " + _
                                               "Mid$(Prd_Data, " + Posicion + ", " + Largo + ")='" + _
                                                PatronBusqueda + "';")
  If FiltroVsdPrd.RecordCount > 0 Then
    ProductoAsociadoEncontrado = True
    DataPatron = FiltroVsdPrd!Prd_Data
  End If
  FiltroVsdPrd.Close
  
End Function





Public Sub CalculaTotales()

Dim TotalApVig, ApVig_Pes, DeAct_Pes, SomAp_Pes As Currency
Dim TotalDeAct, ApVig_Dol, DeAct_Dol, SomAp_Dol As Currency
Dim TotalSomAp, ApVig_UF, DeAct_UF, SomAp_UF As Currency
Dim TotalSug As Long
Dim Garantia As String
Dim Monto1, Monto2 As String

'** calcular total en pesos
ApVig_Pes = 0
ApVig_Dol = 0
ApVig_UF = 0
DeAct_Pes = 0
DeAct_Dol = 0
DeAct_UF = 0
SomAp_Pes = 0
SomAp_Dol = 0
SomAp_UF = 0
TotalSug = 0
Garantia = ""

For j# = 1 To Grid.MaxRows
  
  Grid.Row = j#
  
  Grid.Col = 3
  If Trim$(Grid.Text) <> "" Then
    If Garantia = "" And Trim$(Grid.Text) <> "" Then
      Garantia = Trim$(Grid.Text)
    ElseIf Trim$(Grid.Text) < Garantia And Trim$(Grid.Text) <> "" Then
      Garantia = Trim$(Grid.Text)
    End If
  End If
  
  Grid.Col = 4
  TotalSug = TotalSug + Val(SinPuntos(Grid.Text))
    
  Grid.Col = 5
  If Grid.Text = "$" Then
      Grid.Col = 6
      ApVig_Pes = ApVig_Pes + Val(SinPuntos(Grid.Text))
      Grid.Col = 7
      DeAct_Pes = DeAct_Pes + Val(SinPuntos(Grid.Text))
      Grid.Col = 8
      SomAp_Pes = SomAp_Pes + Val(SinPuntos(Grid.Text))
      
  ElseIf Grid.Text = "US$" Then
      Grid.Col = 6
      ApVig_Dol = ApVig_Dol + CCur(SinPuntos(Grid.Text))
      Grid.Col = 7
      DeAct_Dol = DeAct_Dol + CCur(SinPuntos(Grid.Text))
      Grid.Col = 8
      SomAp_Dol = SomAp_Dol + CCur(SinPuntos(Grid.Text))

  ElseIf Grid.Text = "UF" Then
      Grid.Col = 6
      ApVig_UF = ApVig_UF + CCur(SinPuntos(Grid.Text))
      Grid.Col = 7
      DeAct_UF = DeAct_UF + CCur(SinPuntos(Grid.Text))
      Grid.Col = 8
      SomAp_UF = SomAp_UF + CCur(SinPuntos(Grid.Text))

  End If
  
Next j#

Grid1.Col = 1
Grid1.Text = Garantia

Grid1.Col = 2
Grid1.Text = Format$(TotalSug, "###,###,###,##0")

Grid2.Col = 1
Grid2.Row = 0
Grid2.Text = Format$(ApVig_Pes, "###,###,###,##0")
Grid2.Row = 1
Grid2.Text = Format$(ApVig_Dol, "###,###,###,##0.00")
Grid2.Row = 2
Grid2.Text = Format$(ApVig_UF, "###,###,###,##0.00")

Grid2.Col = 2
Grid2.Row = 0
Grid2.Text = Format$(DeAct_Pes, "###,###,###,##0")
Grid2.Row = 1
Grid2.Text = Format$(DeAct_Dol, "###,###,###,##0.00")
Grid2.Row = 2
Grid2.Text = Format$(DeAct_UF, "###,###,###,##0.00")

Grid2.Col = 3
Grid2.Row = 0
Grid2.Text = Format$(SomAp_Pes, "###,###,###,##0")
Grid2.Row = 1
Grid2.Text = Format$(SomAp_Dol, "###,###,###,##0.00")
Grid2.Row = 2
Grid2.Text = Format$(SomAp_UF, "###,###,###,##0.00")

TotalApVig = ApVig_Pes + (ApVig_Dol * CCur(StrDolar)) + (ApVig_UF * CCur(StrUF))
TotalDeAct = DeAct_Pes + (DeAct_Dol * CCur(StrDolar)) + (DeAct_UF * CCur(StrUF))
TotalSomAp = SomAp_Pes + (SomAp_Dol * CCur(StrDolar)) + (SomAp_UF * CCur(StrUF))

Grid2.Row = 3
Grid2.Col = 1
Grid2.Text = Format$(TotalApVig, "###,###,###,##0")
Grid2.Col = 2
Grid2.Text = Format$(TotalDeAct, "###,###,###,##0")
Grid2.Col = 3
Grid2.Text = Format$(TotalSomAp, "###,###,###,##0")

'** calcula total Riesgo Directo + Riesgo Vinculado
For I% = 1 To 3
  Grid2.Row = 3
  Grid2.Col = I%
  Monto1 = SinPuntos(Grid2.Text)
  Grid2.Row = 4
  Monto2 = SinPuntos(Grid2.Text)
   
  Grid2.Row = 5
  Grid2.Text = Format$(CCur(Monto1) + CCur(Monto2), "###,###,###,##0")
  
Next I%


End Sub

Public Function SinPuntos(Nump As String) As String

SinPuntos = Nump
While InStr(SinPuntos, ".") > 0
  SinPuntos = Mid$(SinPuntos, 1, InStr(SinPuntos, ".") - 1) + Mid$(SinPuntos, InStr(SinPuntos, ".") + 1)
Wend

End Function
