VERSION 4.00
Begin VB.Form ConfiguraProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ModificaciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n Productos a Visar"
   ClientHeight    =   5010
   ClientLeft      =   2445
   ClientTop       =   3225
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Height          =   5415
   Left            =   2385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Top             =   2880
   Width           =   5745
   Begin Threed.SSFrame SSFrame1 
      Height          =   4965
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   5550
      _version        =   65536
      _extentx        =   9790
      _extenty        =   8758
      _stockprops     =   14
      caption         =   "SelecciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n de Productos ..."
      forecolor       =   32768
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
      ' FI2026 - VBX.SpreadSheet migrado a MSFlexGrid (MSFLXGRD.OCX)
      Begin MSFlexGridLib.MSFlexGrid GrillaVigenciaProductos 
         Height          =   3345
         Left            =   150
         TabIndex        =   0
         Top             =   570
         Width           =   5295
         Cols            =   2
         Rows            =   2
         ScrollBars      =   2  'flexScrollBarVertical
         SelectionMode   =   1  'flexSelectionByRow
      End

      Begin Threed.SSCommand Todos 
         Height          =   615
         Left            =   225
         TabIndex        =   1
         Top             =   4170
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Todos"
         forecolor       =   0
         font3d          =   3
         picture         =   "VIGENCIA.frx":0004
      End

      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Estado"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   315
         Width           =   1200
      End

      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Producto"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   1365
         TabIndex        =   5
         Top             =   315
         Width           =   3870
      End

      Begin Threed.SSCommand Aceptar 
         Height          =   615
         Left            =   3555
         TabIndex        =   2
         Top             =   4170
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Aceptar"
         forecolor       =   8421376
         font3d          =   3
         picture         =   "VIGENCIA.frx":0796
      End

      Begin Threed.SSCommand Cancelar 
         Height          =   615
         Left            =   4455
         TabIndex        =   3
         Top             =   4170
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   3
         picture         =   "VIGENCIA.frx":0AB0
      End
   End
End
Attribute VB_Name = "ConfiguraProducto"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim PunteroTabPrd As Recordset
Private Sub Aceptar_Click()
Dim I As Integer
Dim Valor As String
Dim FiltroTabVisar As Recordset
Dim AntiguaVigencia As Integer
Dim CodigoProducto As String
Dim CadenaProducto As String
  I = MsgBox("ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¿Seguro de Cambiar la Vigencia de los Productos?", 4 + 32 + 256, "Confirmar ModificaciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n")
  If I = 6 Then
    Screen.MousePointer = 11
    With GrillaVigenciaProductos
      For I = 1 To .Rows
        .Row = I
        .Col = 4
        AntiguaVigencia = .Text
        .Col = 1
        If AntiguaVigencia <> .Text Then
          If .Text = 1 Then
            Valor = "SI"
          Else
            Valor = "NO"
          End If
          .Col = 3
          CadenaProducto = .Text
          While CadenaProducto <> ""
            CodigoProducto = ObtenerCodigoProducto(CadenaProducto)
            If (ObtieneAliasRegionalDelProducto(CodigoProducto) <> "") And InStr("00,54,55,56", Environ$(NODOMETRO)) = 0 Then
              CodigoProducto = ObtieneAliasRegionalDelProducto(CodigoProducto)
            End If
            If Not EsProductoRelacionado(CodigoProducto) Then
              PunteroTabPrd.Seek "=", Trim(CodigoProducto)
              PunteroTabPrd.Edit
              PunteroTabPrd("Prd_Tag") = Valor
              PunteroTabPrd.Update
              'Set FiltroTabVisar = BaseVisado.CreateSnapshot("SELECT Vis_CodRel FROM " + _
                                                                                               "Tab_Visar WHERE Vis_CodProd='" + _
                                                                                               CodigoProducto + "';")
              'While Not FiltroTabVisar.EOF
              '  If Valor = "NO" Then
              '    If ProductoRelacionadoCompartidoVigente(CodigoProducto, FiltroTabVisar("Vis_CodRel")) Then
              '      Valor = "SI"
              '    End If
              '  End If
              '  PunteroTabPrd.Seek "=", FiltroTabVisar("Vis_CodRel")
              '  PunteroTabPrd.Edit
              '  PunteroTabPrd("Prd_Tag") = Valor
              '  PunteroTabPrd.Update
              '  FiltroTabVisar.MoveNext
              'Wend
              'FiltroTabVisar.Close
            End If
          Wend
        End If
      Next I
    End With
    PunteroTabPrd.Close
    Screen.MousePointer = 0
    Unload Me
  End If
End Sub
Private Sub Aceptar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub
Private Sub Cancelar_Click()
  Unload Me
End Sub

Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub Encabezado_Click(Index As Integer)
  Select Case Index
    Case 0
      Call OrdenarGrilla(1, Index + 1, GrillaVigenciaProductos, 2)
    Case Else
      Call OrdenarGrilla(1, Index + 1, GrillaVigenciaProductos, 1)
  End Select
End Sub
Private Sub Form_Activate()
  Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim CodigoProducto As String
Dim CadenaProducto As String
Dim FiltroTabCfgProd As Recordset
  Screen.MousePointer = 11
  Principal.Enabled = False
  Call CentrarVentana(Me)
  With GrillaVigenciaProductos
    .Cols = 4
    .ColWidth(1) = 10
    .ColWidth(2) = 32
    .ColWidth(3) = 0
    .ColWidth(4) = 0
    ' FI2026 SHG - ColHidden (Spread) no existe en MSFlexGrid. El equivalente es
    ' ColWidth(n) = 0, ya aplicado arriba para las columnas 3 y 4.
    .Col = 3
    '.ColHidden = True
    .Col = 4
    '.ColHidden = True
  End With
  Set FiltroTabCfgProd = BaseVisado.OpenRecordset("SELECT * FROM Tab_CfgProd ORDER BY Cfg_Glosa ASC;")
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  While Not FiltroTabCfgProd.EOF
    CadenaProducto = FiltroTabCfgProd("Cfg_Cod_Prod")
    While CadenaProducto <> ""
      CodigoProducto = ObtenerCodigoProducto(CadenaProducto)
      If (ObtieneAliasRegionalDelProducto(CodigoProducto) <> "") And (InStr("00,54,55,56", Environ$(NODOMETRO)) = 0) Then
        CodigoProducto = ObtieneAliasRegionalDelProducto(CodigoProducto)
      End If
      If Not EsProductoRelacionado(CodigoProducto) Then
        CadenaProducto = ""
        PunteroTabPrd.Seek "=", CodigoProducto
        With GrillaVigenciaProductos
          .Rows = .Rows + 1
          .Row = .Rows
          ' FI2026 SHG - UserResizeRow (Spread) -> AllowUserResizing (MSFlexGrid):
          ' 0=ninguno, 1=columnas, 2=filas, 3=ambos.
          .AllowUserResizing = 2
          .Col = 1
          .BackColor = COLOR_06
          ' FI2026_TODO - CellType=10 + TypeCheckCenter + TypeCheckText (Spread) renderizaba un
          ' CHECKBOX real en la celda con la etiqueta "Vigente"; MSFlexGrid no tiene celda-checkbox
          ' nativa. El toggle 0/1 (mas abajo y en GrillaVigenciaProductos_DblClick) sigue funcionando
          ' sobre .Text, pero visualmente la celda va a mostrar el numero "1"/"0" en vez de un
          ' checkbox marcado/desmarcado. PERDIDA DE FUNCIONALIDAD VISUAL REAL -- revisar con el
          ' usuario si alcanza con texto o si hace falta un CheckBox superpuesto sobre la celda.
          '.CellType = 10
          '.TypeCheckCenter = True
          '.TypeCheckText = "Vigente"
          If PunteroTabPrd("Prd_Tag") = "SI" Then
            .Text = 1
            .Col = 4
            .Text = 1
          Else
            .Text = 0
            .Col = 4
            .Text = 0
          End If
          .Col = 2
          .BackColor = COLOR_06
          .Text = FiltroTabCfgProd("Cfg_Glosa")
          .Col = 3
          .Text = FiltroTabCfgProd("Cfg_Cod_Prod")
        End With
      End If
    Wend
    FiltroTabCfgProd.MoveNext
  Wend
  FiltroTabCfgProd.Close
  Call BloquearEdicionGrilla(1, 2, GrillaVigenciaProductos.Rows, 2, GrillaVigenciaProductos)
  ' FI2026 SHG - CursorType/CursorStyle (Spread) sin equivalente en MSFlexGrid, write-only
  'GrillaVigenciaProductos.CursorType = 0
  'GrillaVigenciaProductos.CursorStyle = 2
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Principal.Enabled = True
End Sub
Private Sub GrillaVigenciaProductos_DblClick()
  ' FI2026 - DblClick(Col,Row) (Spread) -> DblClick() sin parametros en MSFlexGrid
  ' FI2026 - ActiveCol (Spread) -> Col (MSFlexGrid), mismo rol: columna de la celda activa
  If GrillaVigenciaProductos.Col = 2 Then
    GrillaVigenciaProductos.Row = GrillaVigenciaProductos.Row
    GrillaVigenciaProductos.Col = 1
    GrillaVigenciaProductos.Text = Abs(GrillaVigenciaProductos.Text - 1)
  End If
End Sub
Private Sub GrillaVigenciaProductos_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Aceptar_Click
    Case 27
      Unload Me
    Case 32
      Call GrillaVigenciaProductos_DblClick ' FI2026 - sin parametros (ver arriba)
  End Select
End Sub

Private Function ProductoRelacionadoCompartidoVigente(ByVal ProductoPrincipal As String, ByVal ProductoRelacionado As String) As Boolean
Dim FiltroTabVisar As Recordset
  Set FiltroTabVisar = BaseVisado.OpenRecordset("SELECT Vis_CodProd FROM " + _
                                                                                   "Tab_Visar WHERE Vis_CodRel='" + ProductoRelacionado + "';", dbOpenSnapshot)
  ProductoRelacionadoCompartidoVigente = False
  Do While Not FiltroTabVisar.EOF
    If FiltroTabVisar("Vis_CodProd") <> ProductoPrincipal Then
      PunteroTabPrd.Seek "=", FiltroTabVisar("Vis_CodProd")
      If PunteroTabPrd("Prd_Tag") = "SI" Then
        ProductoRelacionadoCompartidoVigente = True
        Exit Do
      End If
    End If
    FiltroTabVisar.MoveNext
  Loop
  FiltroTabVisar.Close
End Function

Private Function ObtenerCodigoProducto(Cadena As String) As String
Dim PosicionComa As Integer
  PosicionComa = InStr(Cadena, ",")
  If PosicionComa = 0 Then
    ObtenerCodigoProducto = Cadena
    Cadena = ""
  Else
    ObtenerCodigoProducto = Mid$(Cadena, 1, PosicionComa - 1)
    Cadena = Mid$(Cadena, PosicionComa + 1)
  End If
End Function
Private Sub Todos_Click()
Dim I As Integer
  With GrillaVigenciaProductos
    .Col = 1
    For I = 1 To .Rows
      .Row = I
      .Text = 1
    Next I
  End With
End Sub
