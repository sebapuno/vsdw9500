VERSION 4.00
Begin VB.Form Visando 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Productos a Visar"
   ClientHeight    =   5190
   ClientLeft      =   4830
   ClientTop       =   2955
   ClientWidth     =   7395
   ClipControls    =   0   'False
   Height          =   5595
   Left            =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Top             =   2610
   Width           =   7515
   Begin Threed.SSPanel Panel3D4 
      Height          =   795
      Left            =   2340
      TabIndex        =   17
      Top             =   4350
      Width           =   1035
      _version        =   65536
      _extentx        =   1826
      _extenty        =   1402
      _stockprops     =   15
      forecolor       =   4194368
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelinner      =   1
      shadowcolor     =   1
      font3d          =   3
      alignment       =   6
      Begin Threed.SSCommand EnviarVisado 
         Height          =   670
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   900
         _version        =   65536
         _extentx        =   1588
         _extenty        =   1191
         _stockprops     =   78
         caption         =   "&Transmitir"
         forecolor       =   32768
         enabled         =   0   'False
         font3d          =   3
         picture         =   "VISANDO.frx":0000
      End
   End

   Begin Threed.SSPanel Panel3D3 
      Height          =   795
      Left            =   3390
      TabIndex        =   14
      Top             =   4350
      Width           =   3870
      _version        =   65536
      _extentx        =   6826
      _extenty        =   1402
      _stockprops     =   15
      forecolor       =   4194368
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelinner      =   1
      shadowcolor     =   1
      font3d          =   3
      alignment       =   6
      Begin Threed.SSCommand Rechazar 
         Height          =   675
         Left            =   1050
         TabIndex        =   22
         Top             =   60
         Width           =   900
         _version        =   65536
         _extentx        =   1588
         _extenty        =   1191
         _stockprops     =   78
         caption         =   "&Borrar"
         forecolor       =   0
         enabled         =   0   'False
         font3d          =   3
         picture         =   "VISANDO.frx":0AE2
      End

      Begin Threed.SSCommand DesMarca 
         Height          =   675
         Left            =   1965
         TabIndex        =   3
         Top             =   60
         Width           =   900
         _version        =   65536
         _extentx        =   1588
         _extenty        =   1191
         _stockprops     =   78
         caption         =   "&Desmarcar"
         forecolor       =   8388608
         enabled         =   0   'False
         font3d          =   3
         picture         =   "VISANDO.frx":1274
      End

      Begin Threed.SSCommand Visar 
         Height          =   675
         Left            =   90
         TabIndex        =   2
         Top             =   60
         Width           =   900
         _version        =   65536
         _extentx        =   1588
         _extenty        =   1191
         _stockprops     =   78
         caption         =   "&Visar"
         forecolor       =   8388608
         font3d          =   3
         picture         =   "VISANDO.frx":1C06
      End

      Begin Threed.SSCommand Cancelar 
         Height          =   675
         Left            =   2895
         TabIndex        =   4
         Top             =   60
         Width           =   900
         _version        =   65536
         _extentx        =   1588
         _extenty        =   1191
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   3
         picture         =   "VISANDO.frx":25EC
      End
   End

   Begin Threed.SSPanel Panel3D2 
      Height          =   3270
      Left            =   90
      TabIndex        =   10
      Top             =   945
      Width           =   7200
      _version        =   65536
      _extentx        =   12700
      _extenty        =   5768
      _stockprops     =   15
      caption         =   "No Existen Datos a Visar para este Cliente"
      forecolor       =   255
      backcolor       =   -2147483633
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   11.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelinner      =   1
      shadowcolor     =   1
      font3d          =   3
      ' FI2026 - VBX.SpreadSheet migrado a MSFlexGrid (MSFLXGRD.OCX)
      Begin MSFlexGridLib.MSFlexGrid GrillaProductos 
         Height          =   2865
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   7010
         Cols            =   4
         Rows            =   13
         ScrollBars      =   2  'flexScrollBarVertical
         SelectionMode   =   1  'flexSelectionByRow
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   2535
         TabIndex        =   13
         Top             =   90
         Width           =   4320
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Fecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   90
         Width           =   960
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Ejecutivo"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   590
         TabIndex        =   11
         Top             =   90
         Width           =   960
      End
   End

   Begin Threed.SSPanel Panel3D1 
      Height          =   735
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   7200
      _version        =   65536
      _extentx        =   12700
      _extenty        =   1296
      _stockprops     =   15
      caption         =   " Datos Cliente"
      forecolor       =   8388608
      backcolor       =   -2147483633
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelinner      =   1
      shadowcolor     =   1
      font3d          =   3
      alignment       =   0
      Begin VB.Label Rentable 
         BackStyle       =   0  'Transparent
         Caption         =   "XX"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6750
         TabIndex        =   23
         Top             =   105
         Width           =   330
      End

      Begin VB.Label NombreCliente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3960
         TabIndex        =   9
         Top             =   345
         Width           =   2955
      End

      Begin VB.Label RutCliente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   585
         TabIndex        =   8
         Top             =   345
         Width           =   2280
      End

      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   1
         Left            =   3240
         TabIndex        =   7
         Top             =   345
         Width           =   585
      End

      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rut"
         BeginProperty Font 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   345
         Width           =   270
      End
   End

   Begin Threed.SSPanel PortaColor 
      Height          =   600
      Left            =   270
      TabIndex        =   19
      Top             =   4395
      Visible         =   0   'False
      Width           =   1545
      _version        =   65536
      _extentx        =   2725
      _extenty        =   1058
      _stockprops     =   15
      backcolor       =   12632256
      borderwidth     =   1
      bevelouter      =   0
      bevelinner      =   2
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rechazado"
         BeginProperty Font 
            name            =   "Courier"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   380
         TabIndex        =   21
         Top             =   330
         Width           =   1150
      End

      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Aprobado"
         BeginProperty Font 
            name            =   "Courier"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   380
         TabIndex        =   20
         Top             =   60
         Width           =   1080
      End

      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   45
         Picture         =   "VISANDO.frx":2AC2
         Top             =   315
         Width           =   285
      End

      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   45
         Picture         =   "VISANDO.frx":32C4
         Top             =   45
         Width           =   290
      End
   End

   Begin VB.Image Aprobado 
      Height          =   240
      Left            =   1890
      Picture         =   "VISANDO.frx":3AC6
      Top             =   4650
      Visible         =   0   'False
      Width           =   285
   End

   Begin VB.Image Rechazado 
      Height          =   240
      Left            =   2205
      Picture         =   "VISANDO.frx":42C8
      Top             =   4650
      Visible         =   0   'False
      Width           =   255
   End

   Begin VB.Label SombraColor 
      BackColor       =   &H00000000&
      Height          =   600
      Left            =   330
      TabIndex        =   18
      Top             =   4455
      Visible         =   0   'False
      Width           =   1545
   End

   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   3270
      Index           =   1
      Left            =   150
      TabIndex        =   16
      Top             =   1020
      Width           =   7200
   End

   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   735
      Index           =   0
      Left            =   150
      TabIndex        =   15
      Top             =   150
      Width           =   7200
   End
End
Attribute VB_Name = "Visando"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Type Estructura
  Rut As String * 10
  Data As String
  Color As Long
  Producto As String * 3
  Correlativo As Integer
End Type

Dim PunteroTabVsd As Recordset
Dim ArregloTemporal() As Estructura
Dim UltimaFilaGrillaProductos As Long ' FI2026 - reemplaza el NewRow de LeaveCell (Spread) para detectar cambio de fila en RowColChange (MSFlexGrid)
Function BuscaRentable(Rut_CO As String) As String
Dim Usuario     As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
Dim Ocurrencias As Integer
Dim j  As Integer
Dim Rut_C As String

    Rut_C = Rut_CO
    BuscaRentable = "XX"
    NodoDestino = DeterminaNodo(AplicacionDestino, "L065")
    If Trim$(NodoDestino) <> SINNODO Then
      Usuario = IdUsuario + "L065" + "L"
      Rut_C = Mid$(DesformateaRut(Rut_C), 1, 9) + "VIS"
      CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Rut_C, "8K")
      If Mid$(CadenaDatos, 1, 6) = "000000" Then
         Ocurrencias = Int(Mid$(CadenaDatos, 8, 4))
         CadenaDatos = Mid$(CadenaDatos, 24)
         For j = 0 To Ocurrencias - 1
            If Mid$(CadenaDatos, 31 + j * 34, 2) = "RN" Then
               BuscaRentable = Mid$(CadenaDatos, 33 + j * 34, 2)
               Exit For
            End If
         Next j
      End If
    End If

End Function

Private Sub LlenarGrillaProductos()
Dim Rut As String
Dim NumeroFila As Integer
Dim FiltroVsdPrd As Recordset
Dim DatoRelevante As String
Dim PunteroTabPrd As Recordset
  Rut = DesformateaRut(RutCliente.Caption)
  Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE Prd_RutCli='" + Rut + _
                                                                                 "' ORDER BY Prd_CodProd, Prd_Correl ASC;", dbOpenSnapshot)
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  NumeroFila = 1
  PunteroTabPrd.Index = "Tab_Prd"
  While Not FiltroVsdPrd.EOF
    If Not EsProductoRelacionado(FiltroVsdPrd("Prd_CodProd")) Then
      PunteroTabPrd.Seek "=", FiltroVsdPrd("Prd_CodProd")
      If NumeroFila = 1 Then
        Call HabilitarGrillaProductos(True)
      End If
      ' FI2026 - era .MaxRows = NumeroFila (Spread cuenta solo filas de datos)
      GrillaProductos.Rows = NumeroFila + 1
      GrillaProductos.Row = NumeroFila
      GrillaProductos.Col = 1
      ' FI2026 - GrillaProductos.TypeHAlign = 2 migrado a GrillaProductos.CellAlignment
      GrillaProductos.CellAlignment = 4  ' flexAlignCenterCenter
      If PunteroTabPrd("Prd_Pos_Eje") <> "0,0" Then
         GrillaProductos.Text = Mid$(FiltroVsdPrd("Prd_Data"), _
                                           Interpreta(PunteroTabPrd("Prd_Pos_Eje"), SIMBOLOPOSICION), _
                                           Interpreta(PunteroTabPrd("Prd_Pos_Eje"), SIMBOLOLARGO))
      Else
         GrillaProductos.Text = "000"
      End If
      GrillaProductos.Col = 2
      ' FI2026 - GrillaProductos.TypeHAlign = 2 migrado a GrillaProductos.CellAlignment
      GrillaProductos.CellAlignment = 4  ' flexAlignCenterCenter
      GrillaProductos.Text = ObtenerDescripcionGlosa(FiltroVsdPrd("Prd_CodProd"), _
                                           Mid$(FiltroVsdPrd("Prd_Data"), _
                                           Interpreta(PunteroTabPrd("Prd_Pos_Fec"), SIMBOLOPOSICION), _
                                           Interpreta(PunteroTabPrd("Prd_Pos_Fec"), SIMBOLOLARGO)), _
                                           PunteroTabPrd("Prd_Pic_Fec"))
      GrillaProductos.Col = 3
      If PunteroTabPrd("Prd_Pos_Dat") <> "0,0" Then
         DatoRelevante = Mid$(FiltroVsdPrd("Prd_Data"), _
                                  Interpreta(PunteroTabPrd("Prd_Pos_Dat"), SIMBOLOPOSICION), _
                                  Interpreta(PunteroTabPrd("Prd_Pos_Dat"), SIMBOLOLARGO))
      Else
         DatoRelevante = ""
      End If
      If Not IsNull(PunteroTabPrd("Prd_Pic_Dat")) Then
        If Trim$(PunteroTabPrd("Prd_Pic_Dat")) <> "" Then
          DatoRelevante = Format$(DatoRelevante, PunteroTabPrd("Prd_Pic_Dat"))
        End If
      End If
      With GrillaProductos
        .Text = PunteroTabPrd("Prd_Descrip") + "  " + DatoRelevante
        .Col = 4
        .Text = FiltroVsdPrd("Prd_CodProd")
        .Col = 5
        .Text = FiltroVsdPrd("Prd_Correl")
        .Col = 6
        .Text = FiltroVsdPrd("Prd_Data")
      End With
      NumeroFila = NumeroFila + 1
    End If
    FiltroVsdPrd.MoveNext
  Wend
  FiltroVsdPrd.Close
  PunteroTabPrd.Close
End Sub

Private Sub Cancelar_Click()
  Unload Me
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub
Private Sub DesMarca_Click()
Dim i As Integer
  With GrillaProductos
    .Row = .Row
    .Col = 1
    If .BackColor <> COLOR_04 Then
      i = MsgBox("¿ Seguro de Desmarcar ?", 4 + 32 + 256, "Verificar Desmarcado")
      If i = 6 Then
        ' FI2026 SHG - OperationMode (Spread) sin equivalente en MSFlexGrid
        '.OperationMode = 3
        For i = 1 To 3
          .Col = i
          .BackColor = COLOR_04
        Next i
        .Col = 0
        ' FI2026 SHG - CellType (Spread) sin equivalente en MSFlexGrid
        '.CellType = 5
        ' FI2026 - .TypeHAlign = 2 migrado a .CellAlignment
        .CellAlignment = 4  ' flexAlignCenterCenter
        ' FI2026 SHG - TypeTextShadow (Spread) sin equivalente en MSFlexGrid
        '.TypeTextShadow = True
        .ForeColor = COLOR_08
        ' FI2026 SHG - OperationMode (Spread) sin equivalente en MSFlexGrid
        '.OperationMode = 2
        If Not QuedanFilasMarcadas() Then
          PortaColor.Visible = False
          SombraColor.Visible = False
          DesMarca.Enabled = False
          EnviarVisado.Enabled = False
        End If
      End If
    End If
    .SetFocus
  End With
End Sub
Private Sub DesMarca_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub EnviarVisado_Click()
Dim Rut As String
Dim Color As Long
Dim Datos As String
Dim Producto As String
Dim TotalFilas As Integer
Dim Correlativo As Integer
Dim FiltroTabVisar As Recordset
Dim FiltroProdVisar As Recordset
Dim PunteroVsdPrd As Recordset
Dim IncrementoAvance As Integer
Dim PatronProductoPrincipal As String
Dim AlMenosUnoTransmitido As Boolean
Dim PatronProductoRelacionado As String
Dim PorLoMenosUnoNoTransmitido As Boolean
Dim SubOpcion As Integer

  EnviarVisado.Enabled = False
  i% = MsgBox("¿ Desea Transmitir Ahora ?", 4 + 32 + 256, "Verificar Transmisión")
  If i% = 6 Then
    Screen.MousePointer = 11
    TotalFilas = 0
    GrillaProductos.Col = 1
    For i% = 1 To GrillaProductos.Rows - 1 ' FI2026 - era .MaxRows
      GrillaProductos.Row = i%
      If GrillaProductos.BackColor <> COLOR_04 Then
        TotalFilas = TotalFilas + 1
      End If
    Next i%
    If TotalFilas > 0 Then
      IncrementoAvance = Int(2615 / TotalFilas)
    End If
    With Principal
      .Avance.Width = 15
      .Avance.Visible = True
      .PanelStatus.Visible = True
      .Encabezado(2).Visible = False
    End With
    BaseVisado.Close
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    Rut = DesformateaRut(RutCliente.Caption)
    TotalFilas = GrillaProductos.Rows - 1 ' FI2026 - era .MaxRows (filas de datos)
    AlMenosUnoTransmitido = False
    PorLoMenosUnoNoTransmitido = False
    i% = 1
    While i% <= TotalFilas
      GrillaProductos.Row = i%
      GrillaProductos.Col = 1
      If GrillaProductos.BackColor <> COLOR_04 Then
        Erase ArregloTemporal()
        Color = GrillaProductos.BackColor
        GrillaProductos.Col = 4
        Producto = GrillaProductos.Text
        GrillaProductos.Col = 5
        Correlativo = GrillaProductos.Text
        PunteroVsdPrd.Seek "=", Rut, Producto, Correlativo
        Datos = PunteroVsdPrd("Prd_Data")
        If Mid$(Producto, 1, 1) <> "1" And Mid$(Producto, 1, 1) <> "2" And Mid$(Producto, 1, 1) <> "3" Then
           Call LlenarArregloTMP(Rut, Producto, Correlativo, Datos, Color)
        End If
        GrillaProductos.Col = 1
        Set FiltroTabVisar = BaseVisado.OpenRecordset("SELECT * FROM Tab_Visar WHERE Vis_CodProd='" + _
                                                                                          Producto + "' ORDER BY Vis_Correl;", dbOpenSnapshot)
        Do While Not FiltroTabVisar.EOF
          If Trim$(FiltroTabVisar("Vis_PosProd")) <> "0,0" Then
            PatronProductoPrincipal = Mid$(Datos, _
                                                                Interpreta(Trim(FiltroTabVisar("Vis_PosProd")), SIMBOLOPOSICION), _
                                                                Interpreta(Trim(FiltroTabVisar("Vis_PosProd")), SIMBOLOLARGO))
            Set FiltroProdVisar = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE Prd_RutCli='" + _
                                                                                               Rut + "' AND Prd_CodProd='" + _
                                                                                               FiltroTabVisar("Vis_CodRel") + "';", dbOpenSnapshot)
            Do While Not FiltroProdVisar.EOF
              If FiltroTabVisar("Vis_PosRel") <> "0,0" Then
                PatronProductoRelacionado = Mid$(FiltroProdVisar("Prd_Data"), _
                                                                           Interpreta(FiltroTabVisar("Vis_PosRel"), SIMBOLOPOSICION), _
                                                                           Interpreta(FiltroTabVisar("Vis_PosRel"), SIMBOLOLARGO))
                If PatronProductoPrincipal = PatronProductoRelacionado Then
                  Call LlenarArregloTMP(Rut, FiltroProdVisar("Prd_CodProd"), _
                                                        FiltroProdVisar("Prd_Correl"), FiltroProdVisar("Prd_Data"), Color)
                End If
              End If
              FiltroProdVisar.MoveNext
            Loop
            FiltroProdVisar.Close
          End If
          FiltroTabVisar.MoveNext
        Loop
        FiltroTabVisar.Close
        If Mid$(Producto, 1, 1) = "1" Or Mid$(Producto, 1, 1) = "2" Or Mid$(Producto, 1, 1) = "3" Then
           Call LlenarArregloTMP(Rut, Producto, Correlativo, Datos, Color)
        End If
        If PrepararTransmision(Producto, Color) Then
          If Mid$(Producto, 1, 1) = "2" Or Mid$(Producto, 1, 1) = "3" Then Call Amortiguar(Rut, Producto, Correlativo, Color)
          If TransmiteVisado Then
            AlMenosUnoTransmitido = True
            Call RealizarGestion(Rut, Producto, Correlativo, Color)
            
            If Color <> COLOR_03 Then
               SubOpcion = 1
            Else
               SubOpcion = 0
            End If
            
            Call LoguearVisado(String(8, " "), SubOpcion, "00000000", Rut, "A", "00", "00", Mid(Datos, 10, 256), Producto)
            
            If GrillaProductos.BackColor = COLOR_03 Then
              Call HacerTransaccionMonetaria(Rut, Producto, Correlativo, Datos)
              Call CrearRespaldoHistorico(Rut, Producto, Correlativo, Datos)
            End If
            BaseVisado.Execute "DELETE * FROM Vsd_Prd WHERE Prd_RutCli='" + Rut + _
                                              "' AND Prd_CodProd='" + Producto + _
                                              "' AND Prd_Correl=" + LTrim$(Str(Correlativo)) + ";"
            i% = i% - 1
            TotalFilas = TotalFilas - 1
            GrillaProductos.RemoveItem GrillaProductos.Row ' FI2026 - Action=5 (Spread, borrar fila) -> RemoveItem (MSFlexGrid)
          Else
            PorLoMenosUnoNoTransmitido = True
            Call RealizarLlamado(Rut, Producto, Correlativo, Color)
            MsgBox "Hubo Errores en la Transmisión del Producto " + Producto, 0 + 48, "Estado de la Transmisión"
          End If
        Else
          PorLoMenosUnoNoTransmitido = True
          MsgBox "Por una Incongruencia en Producto " + Producto + "," + Chr(10) + _
                        "No se ha Realizado Dicho Visado.", 0 + 48, "Estado de la Transmisión"
        End If
        Principal.Avance.Width = Principal.Avance.Width + IncrementoAvance
      End If
      i% = i% + 1
    Wend
    PunteroVsdPrd.Close
    With Principal
      .Avance.Width = 2615
      .Avance.Visible = False
      .Encabezado(2).Visible = True
      .PanelStatus.Visible = False
    End With
    ' FI2026 - era .MaxRows = TotalFilas
    GrillaProductos.Rows = TotalFilas + 1
    
    If Not QuedanFilasMarcadas Then
      If TotalFilas = 0 Then
        Visar.Enabled = False
        Rechazar.Enabled = False ' **
        If Ventana = "RUT" Then
          '*Principal.Grilla.RowHeight(Principal.Grilla.Row) = 0 = False
        Else
          ' FI2026 - era .RowHidden = True (Spread) -> altura 0 en MSFlexGrid; se saca
          ' el '= True' que quedaba pegado (VB6 lo evaluaba como (0 = True) -> 0).
          ' FI2026 - Grilla.Row sin calificar -> Principal.Grilla.Row (el control vive
          ' en VISADO.FRM, no en este form).
          Principal.Grilla.RowHeight(Principal.Grilla.Row) = 0
        End If
      End If
      Me.Refresh
      PortaColor.Visible = False
      SombraColor.Visible = False
      DesMarca.Enabled = False
      EnviarVisado.Enabled = False
    End If
    If AlMenosUnoTransmitido Then
      If PorLoMenosUnoNoTransmitido Then
        MsgBox "Transmisión de Productos Completamente Exitosa" + Chr(10) + "Sólo en un Cierto Porcentaje", 0 + 64, "Estado de la Transmisión"
      Else
        MsgBox "Transmisión de Productos" + Chr(10) + "Completamente Exitosa", 0 + 64, "Estado de la Transmisión"
      End If
    Else
      MsgBox "Transmisión de Productos" + Chr(10) + "Fallida", 0 + 64, "Estado de la Transmisión"
    End If
    Screen.MousePointer = 0
  Else
    EnviarVisado.Enabled = True
  End If
  If TotalFilas <> 0 Then
    'EnviarVisado.Enabled = True
  Else
    Cancelar.SetFocus
  End If
  
End Sub
Private Sub EnviarVisado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
Dim RutCl As String

  Screen.MousePointer = 11
  Principal.Enabled = False
  Call CentrarVentana(Me)
  Call HabilitarGrillaProductos(False)
  
  If Ventana = "RUT" Then
    RutCliente.Caption = GRut_CForma
    
    Rentable.Caption = BuscaRentable(GRut_SForma) '*
    NombreCliente.Caption = TraeNombre(GRut_SForma)
    
    Retorno = PoneValorINI("Parametros", "Param1", GRut_SForma, DIRECTORIO_01 + ARCHIVO_INI01)
    
  Else
    Principal.Grilla.Col = 1 '*
    RutCliente.Caption = Trim$(Principal.Grilla.Text) '*
     
    RutCl = DesformateaRut(RutCliente.Caption)
    
    Retorno = PoneValorINI("Parametros", "Param1", RutCl, DIRECTORIO_01 + ARCHIVO_INI01)
     
    While InStr("123456789", Mid(RutCliente.Caption, 1, 1)) = 0
      RutCliente.Caption = Mid$(RutCliente.Caption, 2)
    Wend
    Rentable.Caption = BuscaRentable(Trim$(Principal.Grilla.Text)) '*
    Principal.Grilla.Col = 2 '*
    NombreCliente.Caption = Principal.Grilla.Text '*
  End If
  
  With GrillaProductos
    ' FI2026 - MaxCols=6 (Spread: columnas de DATOS 1..6) -> Cols=7
    .Cols = 7
    ' FI2026 - el Spread original tenia DisplayColHeaders=0 (sin encabezado de
    ' columna: los titulos son Labels aparte), asi que la fila fija 0 de
    ' MSFlexGrid se oculta. La COLUMNA 0 en cambio queda visible a proposito:
    ' es el encabezado de fila donde ColocarTickGrillaProductos pone el icono
    ' de aprobado/rechazado (.Col = 0 + .CellPicture).
    .RowHeight(0) = 0
    .ColWidth(1) = 960  ' FI2026 - eran 8 caracteres (Spread)
    .ColWidth(2) = 960  ' FI2026 - eran 8 caracteres (Spread)
    .ColWidth(3) = 4320  ' FI2026 - eran 36 caracteres (Spread)
    .ColWidth(4) = 0
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    ' FI2026 SHG - ColHidden (Spread) no existe en MSFlexGrid. El equivalente es
    ' ColWidth(n) = 0, ya aplicado arriba para las columnas 4, 5 y 6.
    .Col = 4
    '.ColHidden = True
    .Col = 5
    '.ColHidden = True
    .Col = 6
    '.ColHidden = True
    ' FI2026 SHG - TypeEditLen/CursorType/CursorStyle (Spread) sin equivalente en MSFlexGrid, write-only
    '.TypeEditLen = 10000
    '.CursorType = 0
    '.CursorStyle = 2
  End With
  Call LlenarGrillaProductos
  ' FI2026 - .Rows - 1: era .MaxRows
  Call BloquearEdicionGrilla(1, 1, GrillaProductos.Rows - 1, 3, GrillaProductos)
  Call OrdenarGrilla(1, 2, GrillaProductos, 1)
  UltimaFilaGrillaProductos = GrillaProductos.Row ' FI2026 - init tracker de fila para GrillaProductos_RowColChange
  
End Sub

Private Sub HabilitarBotonTransmision()
  With Visando
    .DesMarca.Enabled = True
    .EnviarVisado.Enabled = True
    .PortaColor.Visible = True
    .SombraColor.Visible = True
    .GrillaProductos.Row = .GrillaProductos.Row
  End With
End Sub


Private Sub ColocarTickGrillaProductos(ByVal Tick As Image, ByVal ColorFondo As Long)

Visando.GrillaProductos.Row = Visando.GrillaProductos.Row

  With Visando.GrillaProductos
    .Col = 0
    ' FI2026 SHG - CellType (Spread) sin equivalente en MSFlexGrid
    '.CellType = 9
    .BackColor = COLOR_07
    ' FI2026 SHG - TypePictCenter/Stretch/MaintainScale (Spread) sin equivalente en MSFlexGrid; revisar visualmente el ajuste del icono
    '.TypePictCenter = True
    '.TypePictStretch = True
    '.TypePictMaintainScale = True
    ' FI2026_TODO SHG - CellPictureAlignment comentada: no confirmado si existe en esta version de MSFLXGRD.OCX, revisar contra la VM antes de reactivar
    '.CellPictureAlignment = 4 ' FI2026 - flexAlignCenterCenter, reemplaza TypePictCenter/Stretch/MaintainScale
    Set .CellPicture = Tick.Picture ' FI2026 - TypePictPicture (Spread) -> CellPicture (MSFlexGrid)
    For i% = 1 To 3
      .Col = i%
      .BackColor = ColorFondo
    Next i%
  End With
End Sub

Private Function PermitirRechazoProducto() As Boolean
Dim Archivo As String
Dim Retorno As Integer
Dim Producto As String
Dim Correlativo As String
Dim NumeroRut As String
  Screen.MousePointer = 11
  PermitirRechazoProducto = True
  Visando.GrillaProductos.Row = Visando.GrillaProductos.Row
  Visando.GrillaProductos.Col = 4
  Producto = Visando.GrillaProductos.Text
  Archivo = UCase$(DIRECTORIO_02 + Mid(ARCHIVO_EXE06, 1, Len(ARCHIVO_EXE06) - 9) + _
                  Producto + Mid(ARCHIVO_EXE06, Len(ARCHIVO_EXE06) - 5))
                  
  If Dir(Archivo) = Mid$(Archivo, Len(Archivo) - 11) Then
    NumeroRut = DesformateaRut(RutCliente.Caption)
    Visando.GrillaProductos.Col = 5
    Correlativo = Visando.GrillaProductos.Text
    Retorno = PoneValorINI("Parametros", "Param1", NumeroRut, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param2", Producto, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param3", Correlativo, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
    Call EjecutarPrograma(Archivo, 1)
    If SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01) <> "OK" Then
      PermitirRechazoProducto = False
    End If
  End If
  Screen.MousePointer = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
  
  Principal.Enabled = True '*

  If Ventana = "RUT" Then
    Principal.PortaBotones.Visible = True
  Else
    Principal.Grilla.SetFocus  '*
  End If
  
End Sub
Private Sub GrillaProductos_DblClick()
  ' FI2026 - DblClick(Col,Row) (Spread) -> DblClick() sin parametros en MSFlexGrid; Col/Row no se usaban en el cuerpo
  Call Visar_Click
End Sub
Private Sub GrillaProductos_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Visar_Click
    Case 27
      Unload Me
  End Select
End Sub
' FI2026 - LeaveCell(Col,Row,NewCol,NewRow,Cancel) (Spread) no existe en MSFlexGrid.
' RowColChange() dispara despues de cualquier cambio de fila/columna; se compara
' contra UltimaFilaGrillaProductos (fila previa) para reproducir el "If Row <> NewRow"
' original y disparar la misma logica solo cuando cambia de FILA.
Private Sub GrillaProductos_RowColChange()
  If GrillaProductos.Row <> UltimaFilaGrillaProductos Then
    Call GrillaProductos_LostFocus
    UltimaFilaGrillaProductos = GrillaProductos.Row
  End If
End Sub
Private Sub GrillaProductos_LostFocus()
  GrillaProductos.Col = 1
  ' FI2026 SHG - OperationMode (Spread) sin equivalente en MSFlexGrid
  'GrillaProductos.OperationMode = 2
End Sub

Private Sub Rechazar_Click()
  If PermitirRechazoProducto Then
    Call ColocarTickGrillaProductos(Rechazado, COLOR_02)
    Call HabilitarBotonTransmision
  Else
    MsgBox "No se Puede Rechazar Dicho Producto", 0 + 48, "Falla de Rechazo"
  End If

End Sub

Private Sub Rechazar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If

End Sub


Private Sub Titulo_Click(Index As Integer)
  Call OrdenarGrilla(0, Index + 1, GrillaProductos, 1)
End Sub
Private Sub Visar_Click()
Dim Rut           As String
Dim Producto      As String
Dim Correlativo   As Integer
Dim RespVal       As String * 2
Dim Data          As String
Dim Tipo_Producto As String

'Recorte Plataforma Siebel Fase V
'jsa 29/03/2007
'Recorte Visado Cuentas Vista (Cuenta Directa y Cuenta Familiar)
'*******************************************************************
Recorte = False
If EsSiebel Then
   GrillaProductos.Row = GrillaProductos.Row
   GrillaProductos.Col = 4
   Producto = GrillaProductos.Text
   
   If Producto = "2PH" Then
      GrillaProductos.Col = 6
      Data = GrillaProductos.Text
      Tipo_Producto = Mid(Data, 279, 2)
      '10 : Cuenta Directa, 11 : Cuenta Familiar
      If Tipo_Producto = "10" Or Tipo_Producto = "11" Then
          Recorte = True
      End If
    End If
 End If
 If Recorte Then
    MsgBox "Debe Visar Cuentas Vista por Siebel", 64, "Verifica Siebel"
 End If
 '********************************************************************

If Not Recorte Then

  Visar.Enabled = False
  GrillaProductos.Row = GrillaProductos.Row
  GrillaProductos.Col = 1
  If GrillaProductos.BackColor = COLOR_04 Then
    Rut = DesformateaRut(RutCliente.Caption)
    GrillaProductos.Col = 4
    Producto = GrillaProductos.Text
    GrillaProductos.Col = 5
    Correlativo = GrillaProductos.Text
    
    If AprobarProductoUnicaLlave(GrillaProductos.Row) Then
      RespVal = ValidarProducto("1", Rut, Producto, Correlativo)
    End If
    
    If RespVal = "OK" Then
      RespVal = ValidarParametrosCTDLCD(Rut, Producto, Correlativo)
    End If
    
    Select Case RespVal
           Case "EE"
                With Visando.GrillaProductos
                     .Col = 0
                     ' FI2026 SHG - CellType (Spread) sin equivalente en MSFlexGrid
                     '.CellType = 9
                     .BackColor = COLOR_07
                     ' FI2026 SHG - TypePictCenter/Stretch/MaintainScale (Spread) sin equivalente en MSFlexGrid; revisar visualmente el ajuste del icono
                     '.TypePictCenter = True
                     '.TypePictStretch = True
                     '.TypePictMaintainScale = True
                     ' FI2026_TODO SHG - CellPictureAlignment comentada: no confirmado si existe en esta version de MSFLXGRD.OCX, revisar contra la VM antes de reactivar
                     '.CellPictureAlignment = 4 ' FI2026 - flexAlignCenterCenter, reemplaza TypePictCenter/Stretch/MaintainScale
                     Set .CellPicture = Aprobado.Picture ' FI2026 - TypePictPicture (Spread) -> CellPicture (MSFlexGrid)
                     For i% = 1 To 3
                       .Col = i%
                       .BackColor = COLOR_03
                     Next i%
                End With
                With Visando
                     .DesMarca.Enabled = True
                     .EnviarVisado.Enabled = True
                     .PortaColor.Visible = True
                     .SombraColor.Visible = True
                     .GrillaProductos.Row = .GrillaProductos.Row
                     .Panel3D3.Enabled = True
                     .Panel3D4.Enabled = True
                     .Enabled = True
                     If .GrillaProductos.BackColor <> COLOR_04 Then
                        .GrillaProductos.Col = 1
                        ' FI2026 SHG - OperationMode (Spread) sin equivalente en MSFlexGrid
                        '.GrillaProductos.OperationMode = 0
                     End If
                     .GrillaProductos.SetFocus
                 End With
           Case "RR"
                With Visando.GrillaProductos
                     .Col = 0
                     ' FI2026 SHG - CellType (Spread) sin equivalente en MSFlexGrid
                     '.CellType = 9
                     .BackColor = COLOR_07
                     ' FI2026 SHG - TypePictCenter/Stretch/MaintainScale (Spread) sin equivalente en MSFlexGrid; revisar visualmente el ajuste del icono
                     '.TypePictCenter = True
                     '.TypePictStretch = True
                     '.TypePictMaintainScale = True
                     ' FI2026_TODO SHG - CellPictureAlignment comentada: no confirmado si existe en esta version de MSFLXGRD.OCX, revisar contra la VM antes de reactivar
                     '.CellPictureAlignment = 4 ' FI2026 - flexAlignCenterCenter, reemplaza TypePictCenter/Stretch/MaintainScale
                     Set .CellPicture = Rechazado.Picture ' FI2026 - TypePictPicture (Spread) -> CellPicture (MSFlexGrid)
                     For i% = 1 To 3
                       .Col = i%
                       .BackColor = COLOR_02
                     Next i%
                End With
                With Visando
                     .DesMarca.Enabled = True
                     .EnviarVisado.Enabled = True
                     .PortaColor.Visible = True
                     .SombraColor.Visible = True
                     .GrillaProductos.Row = .GrillaProductos.Row
                     .Panel3D3.Enabled = True
                     .Panel3D4.Enabled = True
                     .Enabled = True
                     If .GrillaProductos.BackColor <> COLOR_04 Then
                        .GrillaProductos.Col = 1
                        ' FI2026 SHG - OperationMode (Spread) sin equivalente en MSFlexGrid
                        '.GrillaProductos.OperationMode = 0
                     End If
                     .GrillaProductos.SetFocus
                 End With
           Case "OK"
                 PantallaVisar.Show
               
    End Select
  Else
    MsgBox "Producto esta Listo para Transmitir" + Chr(13) + "Si Desea Utilice la Opción Desmarcar", 0 + 16, "Re-Intento de Visado"
    GrillaProductos.SetFocus
  End If
  Visar.Enabled = True
  Rechazar.Enabled = True
End If

End Sub
Private Sub Visar_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Visar_Click
    Case 27
      Unload Me
  End Select
End Sub
Private Sub HabilitarGrillaProductos(ByVal Estado As Boolean)
  Visar.Enabled = Estado
  Rechazar.Enabled = Estado
  GrillaProductos.Visible = Estado
  For i% = 0 To 2
    Titulo(i%).Visible = Estado
  Next i%
  
End Sub
Private Function PrepararTransmision(ByVal Producto As String, ByVal Color As Long) As Boolean
Dim i As Integer
Dim Id As String
Dim Data As String
Dim Llave As String
Dim Archivo As String
Dim Borrado As String
Dim Actividad As String * 1
Dim ArchivoSRM As Database
Dim PunteroTabPrd As Recordset
Dim PunteroTabOper As Recordset
Dim ResultadoBorrado As Boolean
  Set ArchivoSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  ArchivoSRM.Execute "DELETE * FROM Mensaje"
  ArchivoSRM.Close
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  Set PunteroTabVsd = BaseVisado.OpenRecordset("Tab_Vsd", dbOpenTable)
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  ' FI2026 - PunteroTabVsd.Index = PunteroTabVsd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  ' FI2026 - PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PrepararTransmision = False
  PunteroTabPrd.Seek "=", Producto
  Call GrabarTxSAT(PunteroTabPrd("Prd_Transac"), Color, PunteroTabPrd("Prd_Sincro"), Producto)
  For i = 0 To UBound(ArregloTemporal)
    PunteroTabPrd.Seek "=", ArregloTemporal(i).Producto
    If Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOPOSICION) > 0 Then
      Actividad = Mid$(ArregloTemporal(i).Data, _
                                 Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOPOSICION), _
                                 Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOLARGO))
    Else
      Actividad = Mid$(PunteroTabPrd("Prd_Pos_Act"), 1, 1)
    End If
    PunteroTabOper.Seek "=", ArregloTemporal(i).Producto, Actividad
    If Not PunteroTabOper.NoMatch Then
      If (ArregloTemporal(i).Color = COLOR_03) Or (PunteroTabPrd("Prd_EnviaRez")) Then
        Archivo = UCase(DIRECTORIO_02 + _
                                     Mid$(ARCHIVO_EXE03, 1, Len(ARCHIVO_EXE03) - 9) + _
                                     ArregloTemporal(i).Producto + "00" + _
                                     Mid$(ARCHIVO_EXE03, Len(ARCHIVO_EXE03) - 3, Len(ARCHIVO_EXE03)))
        If Dir(Archivo) <> Mid$(Archivo, Len(Archivo) - 11) Then
          Id = ""
          Data = ""
          Llave = "NULL"
          If (PunteroTabPrd("Prd_Pos_DataDst") <> "0,0") And (PunteroTabPrd("Prd_Pos_LlaveDst") <> "0,0") Then
            Id = IdUsuario + PunteroTabPrd("Prd_Vista_Dst") + Trim$(PunteroTabOper("Oper_Srm"))
            If Interpreta(PunteroTabPrd("Prd_Pos_DataDst"), SIMBOLOLARGO) > 0 Then
              Data = Mid$(ArregloTemporal(i).Data, _
                                   Interpreta(PunteroTabPrd("Prd_Pos_DataDst"), SIMBOLOPOSICION), _
                                   Interpreta(PunteroTabPrd("Prd_Pos_DataDst"), SIMBOLOLARGO))
            End If
            Llave = Mid$(ArregloTemporal(i).Data, _
                                  Interpreta(PunteroTabPrd("Prd_Pos_LlaveDst"), SIMBOLOPOSICION), _
                                  Interpreta(PunteroTabPrd("Prd_Pos_LlaveDst"), SIMBOLOLARGO))
          End If
          Call GrabarVisadoEnEnvioSRM(Id, Llave, Data)
          ResultadoBorrado = BorrarEnVistaOrigen(PunteroTabPrd("Prd_Vista_Ori"), ArregloTemporal(i).Data, _
                                                                               Trim(PunteroTabOper("Oper_Borrado")))
          PrepararTransmision = ResultadoBorrado
        Else
          Retorno% = PoneValorINI("Parametros", "Param1", _
                                                    ArregloTemporal(i).Rut, DIRECTORIO_01 + ARCHIVO_INI01)
          Retorno% = PoneValorINI("Parametros", "Param2", _
                                                    ArregloTemporal(i).Producto, DIRECTORIO_01 + ARCHIVO_INI01)
          Retorno% = PoneValorINI("Parametros", "Param3", _
                                                    Format$(LTrim(Str(ArregloTemporal(i).Correlativo)), "00"), _
                                                    DIRECTORIO_01 + ARCHIVO_INI01)
          If ArregloTemporal(i).Color = COLOR_03 Then
             Retorno% = PoneValorINI("Parametros", "Param4", "1", DIRECTORIO_01 + ARCHIVO_INI01)
          Else
             Retorno% = PoneValorINI("Parametros", "Param4", "0", DIRECTORIO_01 + ARCHIVO_INI01)
          End If
          Retorno% = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
          Call EjecutarPrograma(Archivo, 1)
          If SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01) = "OK" Then
            'ResultadoBorrado = BorrarEnVistaOrigen(PunteroTabPrd("Prd_Vista_Ori"), ArregloTemporal(I).Data, _
                                                                                 Trim(PunteroTabOper("Oper_Borrado")))
            PrepararTransmision = True
          End If
        End If
      Else
        ResultadoBorrado = BorrarEnVistaOrigen(PunteroTabPrd("Prd_Vista_Ori"), ArregloTemporal(i).Data, _
                                                                             Trim(PunteroTabOper("Oper_Borrado")))
        PrepararTransmision = ResultadoBorrado
      End If
    Else
      MsgBox "Producto " + Chr(34) + ArregloTemporal(i).Producto + Chr(34) + " No Tiene Operación" + Chr(10) + _
                    "SRM en Tabla " + Chr(34) + "Tab_Oper" + Chr(34) + "," + Chr(10) + "Producto No Será Visado", _
                    0 + 48, "Falta Operación"
    End If
  Next i
  PunteroTabPrd.Close
  PunteroTabVsd.Close
  PunteroTabOper.Close
End Function
Private Sub GrabarTxSAT(ByVal Transaccion As String, ByVal Color As Long, ByVal EsSincronico As Boolean, ByVal ProductoPrincipal As String)
Dim ArchivoSRM As Database
Dim PunteroSRM As Recordset
  Set ArchivoSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroSRM = ArchivoSRM.OpenRecordset("Mensaje")
  If (ObtieneCodigoOriginalProducto(ProductoPrincipal) <> "") And InStr("00,54,55,56", Environ$(NODOMETRO)) = 0 Then
    ProductoPrincipal = ObtieneCodigoOriginalProducto(ProductoPrincipal)
  End If
  PunteroSRM.AddNew
  If (Transaccion <> "XXXX") And (Color <> COLOR_03) Then Transaccion = "1" + Mid$(Transaccion, 2, 3)
  If Not Transaccion = "XXXX" Then
     PunteroSRM("Sat") = True
     PunteroSRM("Tx_Sat") = Transaccion
     PunteroSRM("Tx_Sincro") = EsSincronico
     PunteroSRM("Cod_ProdPrinc") = ProductoPrincipal
  Else
     PunteroSRM("Sat") = False
     PunteroSRM("Tx_Sat") = "XXXX"
     PunteroSRM("Tx_Sincro") = False
     PunteroSRM("Cod_ProdPrinc") = ProductoPrincipal
  End If
  PunteroSRM.Update
  PunteroSRM.Close
  ArchivoSRM.Close
End Sub
Private Function ObtieneCodigoOriginalProducto(ByVal Producto As String) As String
Dim Criterio As String
Dim PunteroTabAlias As Recordset
  Set PunteroTabAlias = BaseVisado.OpenRecordset("Tab_Alias", dbOpenDynaset)
  Criterio = "Alias_Prod_Nuev='" + Producto + "'"
  ObtieneCodigoOriginalProducto = ""
  PunteroTabAlias.FindFirst Criterio
  If Not PunteroTabAlias.NoMatch Then
    ObtieneCodigoOriginalProducto = PunteroTabAlias("Alias_Prod_Orig")
  End If
  PunteroTabAlias.Close
End Function
Private Function GrabarVisadoEnEnvioSRM(ByVal Id As String, ByVal Llave As String, ByVal Data As String)
Dim ArchivoSRM As Database
Dim PunteroSRM As Recordset
  Set ArchivoSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroSRM = ArchivoSRM.OpenRecordset("Mensaje")
  PunteroSRM.MoveLast
  PunteroSRM.Edit
  If IsNull(PunteroSRM("Data")) Then
     PunteroSRM("Data") = Id + Llave + Data + "##"
  Else
     PunteroSRM("Data") = PunteroSRM("Data") + Id + Llave + Data + "##"
  End If
  PunteroSRM.Update
  PunteroSRM.Close
  ArchivoSRM.Close
End Function
Private Function ValidarProducto(ByVal ModoValidacion As String, ByVal RutCliente As String, ByVal CodigoProducto As String, ByVal Correlativo As Integer) As String
Dim Retorno As Integer
Dim NombreArchivo As String
Dim PunteroVsdPrd As Recordset

  Screen.MousePointer = 11
  Retorno = PoneValorINI("PRODUCTO", "Retorno", "", DIRECTORIO_01 + ARCHIVO_INI01)
  Retorno = PoneValorINI("PRODUCTO", "Oficina", "", DIRECTORIO_01 + ARCHIVO_INI01)
  Retorno = PoneValorINI("PRODUCTO", "Region", "", DIRECTORIO_01 + ARCHIVO_INI01)
  Retorno = PoneValorINI("Parametros", "Param1", RutCliente, DIRECTORIO_01 + ARCHIVO_INI01)
 
  NombreArchivo = UCase(DIRECTORIO_02 + _
                                Mid$(ARCHIVO_EXE01, 1, Len(ARCHIVO_EXE01) - 7) + _
                                CodigoProducto + _
                                Mid$(ARCHIVO_EXE01, Len(ARCHIVO_EXE01) - 3, Len(ARCHIVO_EXE01)))
  If Dir(NombreArchivo) = Mid$(NombreArchivo, Len(NombreArchivo) - 11) Then
    BaseVisado.Close
    Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param1", RutCliente, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param2", CodigoProducto, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param3", LTrim$(Str(Correlativo)), DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param4", ModoValidacion, DIRECTORIO_01 + ARCHIVO_INI01)
    Call EjecutarPrograma(NombreArchivo, 1)
    ValidarProducto = SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01)
    ValidarProducto = StrConv(ValidarProducto, 1)
    
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    
    Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    PunteroVsdPrd.Seek "=", RutCliente, CodigoProducto, Correlativo
    Visando.GrillaProductos.Row = Visando.GrillaProductos.Row
    Visando.GrillaProductos.Col = 6
    Visando.GrillaProductos.Text = PunteroVsdPrd("Prd_Data")
    PunteroVsdPrd.Close
    
  Else
    ValidarProducto = "OK"
  End If
  Screen.MousePointer = 0
End Function
Private Function TransmiteVisado() As Boolean
  TransmiteVisado = False
  Retorno% = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
  Call EjecutarPrograma(DIRECTORIO_02 + ARCHIVO_EXE04, 1)
  If SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01) = "OK" Then
    TransmiteVisado = True
  End If
End Function
Private Sub HacerTransaccionMonetaria(ByVal Rut As String, ByVal Producto As String, ByVal Correlativo As Integer, ByVal DataTemporal As String)
Dim Uso As String
Dim Datos As String
Dim Monto As String
Dim Cuenta As String
Dim PunteroTabTX As Recordset
  Set PunteroTabTX = BaseVisado.OpenRecordset("Tab_Tx", dbOpenTable)
  ' FI2026 - PunteroTabTX.Index = PunteroTabTX.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabTX.Seek "=", Producto
  If Not PunteroTabTX.NoMatch Then
    If PunteroTabTX("Tx_Prod") = Producto Then
      Uso = PunteroTabTX("Tx_Uso")
      Cuenta = PunteroTabTX("Tx_Cuenta")
      If (Val(Mid(DataTemporal, Interpreta(Cuenta, SIMBOLOPOSICION), Interpreta(Cuenta, SIMBOLOLARGO))) <> 0) _
        And (Mid$(DataTemporal, Interpreta(Uso, SIMBOLOPOSICION), Interpreta(Uso, SIMBOLOLARGO)) = "N") Then
        Monto = PunteroTabTX("Tx_Monto")
        Datos = PunteroTabTX("Tx_Cod") + "D0" + _
                      Mid$(DataTemporal, Interpreta(Cuenta, SIMBOLOPOSICION), Interpreta(Cuenta, SIMBOLOLARGO)) + _
                      "^70^" + _
                      Mid$(DataTemporal, Interpreta(Monto, SIMBOLOPOSICION), Interpreta(Monto, SIMBOLOLARGO)) + _
                      String(234, "|") + _
                      Mid$(DataTemporal, Interpreta(Monto, SIMBOLOPOSICION), Interpreta(Monto, SIMBOLOLARGO)) + _
                      IdUsuario + ObtenerDigitoVerificador(IdUsuario) + "NN" + String(170, " ") + "A"
        Open DIRECTORIO_03 + ARCHIVO_TXT01 For Output As #1
        Print #1, Datos
        Close #1
        Call EjecutarPrograma(DIRECTORIO_02 + PunteroTabTX("Tx_Prg"), 1)
        Open DIRECTORIO_03 + ARCHIVO_TXT01 For Input As #1
        Datos = Input$(462, #1)
        Close #1
        If Mid$(Datos, 292, 4) = "9999" Then
           MsgBox "Transacción Monetaria Pendiente", 0 + 48, "Atención"
        End If
        Call LoguearTransaccion("00000000", Datos)
      End If
    End If
  End If
  PunteroTabTX.Close
End Sub
Private Sub CrearRespaldoHistorico(ByVal Rut As String, ByVal Producto As String, ByVal Correlativo As Integer, ByVal Datos As String)
Dim PunteroDBF As Recordset
Dim BaseDbfCRE As Database
Dim PunteroVsdNot As Recordset
Dim PunteroVsdHst As Recordset
  Set PunteroVsdHst = BaseVisado.OpenRecordset("Vsd_Hst", dbOpenTable)
  PunteroVsdHst.AddNew
  PunteroVsdHst("Hst_Rut") = Rut
  PunteroVsdHst("Hst_Tipo") = Producto
  PunteroVsdHst("Hst_Data") = Datos
  PunteroVsdHst("Hst_Fech") = Format$(Now, "yyyymmdd")
  PunteroVsdHst("Hst_Correl") = Correlativo
  PunteroVsdHst.Update
  PunteroVsdHst.Close
  If Producto = "CRE" Then
    Set PunteroVsdNot = BaseVisado.OpenRecordset("Vsd_PagNot", dbOpenTable)
    ' FI2026 - PunteroVsdNot.Index = PunteroVsdNot.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    PunteroVsdNot.Seek "=", Rut, Mid$(Datos, 20, 6), Mid$(Datos, 27, 3)
    If Not PunteroVsdNot.NoMatch Then
      Set BaseDbfCRE = OpenDatabase(DIRECTORIO_01, True, False, "dbase III;")
      Set PunteroDBF = BaseDbfCRE.OpenRecordset(TABLA_DBF01, dbOpenTable)
      PunteroDBF.AddNew
      PunteroDBF("T_O") = PunteroVsdNot("T_O")
      PunteroDBF("Rut_Cli") = Rut
      PunteroDBF("Nro_Pag") = PunteroVsdNot("Nro_Pag")
      PunteroDBF("Nom_Not") = PunteroVsdNot("Nom_Not")
      PunteroDBF("Valor_Not") = PunteroVsdNot("Valor_Not")
      PunteroDBF("Valor_Pag") = PunteroVsdNot("Valor_Pag")
      PunteroDBF("Fecha_Vis") = PunteroVsdNot("Fecha_Vis")
      PunteroDBF.Update
      PunteroDBF.Close
      BaseDbfCRE.Close
      PunteroVsdNot.Delete
    End If
    PunteroVsdNot.Close
  End If
End Sub
Sub LlenarArregloTMP(ByVal Rut As String, ByVal Producto As String, ByVal Correlativo As Integer, ByVal Datos As String, ByVal ColorFila As Long)
Dim Indice As Integer
On Error GoTo InicializaIndice
  Indice = UBound(ArregloTemporal) + 1
On Error GoTo 0
  ReDim Preserve ArregloTemporal(Indice)
  With ArregloTemporal(Indice)
    .Rut = Rut
    .Data = Datos
    .Color = ColorFila
    .Producto = Producto
    .Correlativo = Correlativo
  End With
  Exit Sub
  
InicializaIndice:
  Indice = 0
  Resume Next
End Sub
Private Function QuedanFilasMarcadas() As Boolean
Dim i As Integer
  QuedanFilasMarcadas = False
  With GrillaProductos
    .Col = 1
    For i = 1 To .Rows - 1 ' FI2026 - era .MaxRows
      .Row = i
      If .BackColor <> COLOR_04 Then
        QuedanFilasMarcadas = True
        Exit For
      End If
    Next i
  End With
End Function
Private Function BorrarEnVistaOrigen(ByVal VistaOrigen As String, ByVal Data As String, ByVal OperacionDeEliminacion As String) As Boolean
Dim Id As String
Dim DataBorrado As String
Dim StrEstado As String * 1
Dim p_VsdPrd As Recordset
Dim PunteroVistas As Recordset

  BorrarEnVistaOrigen = True
'  StrEstado = SacaValorINI("Parametros", "Param4", DIRECTORIO_01 + ARCHIVO_INI01)
'  If StrEstado = "1" Or StrEstado = "N" Then
  If (ArregloTemporal(i).Color = COLOR_03) Then
     StrEstado = "A"
  Else
     StrEstado = "R"
  End If
  If VistaOrigen <> "XXXX" Then
    If Mid(VistaOrigen, 1, 2) = "SY" Then
       If Not Es_plt_per01(VistaOrigen) Then '*If Val(Mid(VistaOrigen, 3, 2)) > 12 Then
          Id = "b12256781234W0000001Nplt_emp01 exec modifica_estado_empresa "
          DataBorrado = "'" + Mid$(Data, 1, 3) + "'," + "'" + Mid$(Data, 4, 3) + "'," _
                   + "'" + Mid$(Data, 7, 3) + "'," + "'" + Mid$(Data, 10, 30) + "'," _
                   + "'" + Mid$(Data, 40, 1) + "'," + "'" + Mid$(Data, 41, 2) + "'," _
                   + "'" + Mid$(Data, 43, 8) + "'," + "'" + Mid$(Data, 51, 10) + "'"
       Else
          'Id = "b12256781234W0000001Nplt_per01 exec modifica_estado "
          Id = "b12256781234W0000001Nplt_per01 exec vsd_modifica_estado "
          DataBorrado = "'" + Mid$(Data, 1, 3) + "'," + "'" + Mid$(Data, 4, 3) + "'," _
                   + "'" + Mid$(Data, 7, 3) + "'," + "'" + Mid$(Data, 10, 30) + "'," _
                   + "'" + Mid$(Data, 40, 1) + "'," + "'" + Mid$(Data, 41, 2) + "'," _
                   + "'" + Mid$(Data, 43, 8) + "'," + "'" + Mid$(Data, 51, 10) + "'," _
                   + "'" + StrEstado + "'"
       End If
       Call GrabarVisadoEnEnvioSRM(Id, DataBorrado, "")
    Else
       '** Encuentra verdadera vista origen
'Desarrollado  : Entersoft
'fecha         : 02-10-2009
'Observaciones : Se modifica Select de borrado de registros en base local Visado.mdb, tabla Vsd_Prd por estar mal formateado, generando mensaje
'                de error "Run-time error '3075: Syntax error en la expresion de la consulta"
'Scipt original  ***
     'Set p_VsdPrd = BaseVisado.OpenRecordset("Select * From Vsd_Prd Where Mid$(PRD_DATA,1,40) = '" + Mid$(Data, 1, 40) + "'", dbOpenDynaset)
'Nuevo script    ***
      Set p_VsdPrd = BaseVisado.OpenRecordset("Select * From Vsd_Prd Where (Mid$(PRD_DATA, 1, 40)) = '(Mid$(Data, 1, 40))'", dbOpenDynaset)
'************************
       If p_VsdPrd.RecordCount > 0 Then
         If Trim$(p_VsdPrd!PRD_VTAORI) <> "" Then
           VistaOrigen = p_VsdPrd!PRD_VTAORI
         End If
       End If
       p_VsdPrd.Close
       '***
       
       '** Encuentra Operación Borrado Origen
       Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
       Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
       PunteroVistas.Index = "Vistas"
       PunteroVistas.Seek "=", VistaOrigen
       If Not PunteroVistas.EOF Then
          If Trim$(PunteroVistas!Vta_OperBorr) <> "" Then
            OperacionDeEliminacion = PunteroVistas!Vta_OperBorr
          End If
       End If
       PunteroVistas.Close
       BaseEnvioSRM.Close
       '**
       
       PunteroTabVsd.Seek "=", VistaOrigen
       If Not PunteroTabVsd.NoMatch Then
         Id = IdUsuario + VistaOrigen + OperacionDeEliminacion
         DataBorrado = Mid$(Data, 1, PunteroTabVsd("Vsd_L_LlaveOri"))
         Call GrabarVisadoEnEnvioSRM(Id, DataBorrado, "")
       Else
         BorrarEnVistaOrigen = False
       End If
    End If
  End If
End Function
Private Sub RealizarGestion(ByVal Rut As String, ByVal Producto As String, ByVal Correlativo As Integer, ByVal ColorFila As Long)
Dim Retorno As Integer
Dim NombreArchivo As String
  NombreArchivo = UCase(DIRECTORIO_02 + Mid$(ARCHIVO_EXE08, 1, 4) + _
                               Producto + Mid$(ARCHIVO_EXE08, 8))
  If Dir(NombreArchivo) = Mid$(NombreArchivo, Len(NombreArchivo) - 10) Then
    Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param1", Rut, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param2", Producto, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param3", LTrim$(Str(Correlativo)), DIRECTORIO_01 + ARCHIVO_INI01)
    If ColorFila = COLOR_03 Then
      Retorno = PoneValorINI("Parametros", "Param4", "NULL", DIRECTORIO_01 + ARCHIVO_INI01)
    Else
      Retorno = PoneValorINI("Parametros", "Param4", "908", DIRECTORIO_01 + ARCHIVO_INI01)
    End If
    Call EjecutarPrograma(NombreArchivo, 1)
  End If
End Sub
Private Sub RealizarLlamado(ByVal Rut As String, ByVal Producto As String, ByVal Correlativo As Integer, ByVal ColorFila As Long)
Dim Retorno As Integer
Dim NombreArchivo As String
  NombreArchivo = UCase(DIRECTORIO_02 + Mid$(ARCHIVO_EXE11, 1, 4) + _
                               Producto + Mid$(ARCHIVO_EXE11, 8))
  If Dir(NombreArchivo) = Mid$(NombreArchivo, Len(NombreArchivo) - 10) Then
    Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param1", Rut, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param2", Producto, DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno = PoneValorINI("Parametros", "Param3", LTrim$(Str(Correlativo)), DIRECTORIO_01 + ARCHIVO_INI01)
    If ColorFila = COLOR_03 Then
      Retorno = PoneValorINI("Parametros", "Param4", "NULL", DIRECTORIO_01 + ARCHIVO_INI01)
    Else
      Retorno = PoneValorINI("Parametros", "Param4", "908", DIRECTORIO_01 + ARCHIVO_INI01)
    End If
    Call EjecutarPrograma(NombreArchivo, 1)
  End If
End Sub

Private Sub Amortiguar(ByVal Rut As String, ByVal Producto As String, ByVal Correlativo As Integer, ByVal ColorFila As Long)
Dim Retorno As Integer
Dim NombreArchivo As String
  NombreArchivo = UCase(DIRECTORIO_02 + ARCHIVO_EXE10)
  If ColorFila = COLOR_03 Then 'aprueba visado
     If Dir(NombreArchivo) = Mid$(NombreArchivo, Len(NombreArchivo) - 11) Then
        Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "Param1", Rut, DIRECTORIO_01 + ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "Param2", Producto, DIRECTORIO_01 + ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "Param3", LTrim$(Str(Correlativo)), DIRECTORIO_01 + ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "Param4", "NULL", DIRECTORIO_01 + ARCHIVO_INI01)
        Call EjecutarPrograma(NombreArchivo, 1)
     End If
  End If
End Sub


Public Function Es_plt_per01(c_vista As String) As Boolean
Dim PunteroVista As Recordset

Set PunteroVista = BaseVisado.OpenRecordset("Select * From tab_vsd Where VSD_VISTA_ORI='" + c_vista + "'", dbOpenDynaset)
If PunteroVista.RecordCount > 0 Then
   If InStr(UCase(PunteroVista!vsd_operacion), "PLT_PER01 ") > 0 Then
      Es_plt_per01 = True
   Else
      Es_plt_per01 = False
   End If
End If
PunteroVista.Close

End Function

Public Function TraeNombre(iRutC As String) As String

Dim PunteroVsdCli As Recordset

Set PunteroVsdCli = BaseVisado.OpenRecordset("Vsd_Cli", dbOpenDynaset)
PunteroVsdCli.FindFirst "Cli_Rut='" + iRutC + "'"
If PunteroVsdCli.RecordCount = 0 Then
  TraeNombre = Mid$(ConsultaL001(iRutC), 12, 45)
Else
  If Trim$(PunteroVsdCli("Cli_Nombre")) = "" Then
    TraeNombre = Mid$(ConsultaL001(iRutC), 12, 45)
  Else
    TraeNombre = PunteroVsdCli("Cli_Nombre")
  End If
End If
TraeNombre = StrConv(SacarEspaciosIntermedios(TraeNombre), 3)

PunteroVsdCli.Close

End Function

Public Function ConsultaL001(rRutC As String) As String

Dim AplicacionDestino As String * 4
 
ConsultaL001 = ""
NodoDestino = DeterminaNodo(AplicacionDestino, "L001")
If Trim$(NodoDestino) <> SINNODO Then
   Usuario = IdUsuario + "L001" + "L"
   CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, rRutC, "8K")
   If Mid$(CadenaDatos, 1, 6) = "000000" Then
     ConsultaL001 = Mid$(CadenaDatos, 7, 200)
   End If
End If

End Function

Public Function ValidarParametrosCTDLCD(RutCliente As String, CodigoProducto As String, Correlativo As Integer) As String

Dim Retorno As Integer
Dim NombreArchivo As String
Dim PunteroVsdPrd As Recordset

Screen.MousePointer = 11

NombreArchivo = UCase(DIRECTORIO_02 + ARCHIVO_VLCTD)

If Dir(NombreArchivo) = Mid$(NombreArchivo, Len(NombreArchivo) - 11) Then
  BaseVisado.Close
   
  Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
  Retorno = PoneValorINI("Parametros", "Param1", RutCliente, DIRECTORIO_01 + ARCHIVO_INI01)
  Retorno = PoneValorINI("Parametros", "Param2", CodigoProducto, DIRECTORIO_01 + ARCHIVO_INI01)
  Retorno = PoneValorINI("Parametros", "Param3", LTrim$(Str(Correlativo)), DIRECTORIO_01 + ARCHIVO_INI01)
   
  Call EjecutarPrograma(NombreArchivo, 1)
  
  ValidarParametrosCTDLCD = SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01)
  ValidarParametrosCTDLCD = StrConv(ValidarParametrosCTDLCD, 1)
  
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  
Else
    ValidarParametrosCTDLCD = "OK"
End If

Screen.MousePointer = 0

End Function

Public Function AprobarProductoUnicaLlave(Fila As Integer) As Boolean

Dim NRow1%, NRow2%
Dim ProdFila1 As String
Dim LlaveFila1 As String
Dim FechaFila1 As String
Dim ProdFila2 As String
Dim LlaveFila2 As String
Dim FechaFila2 As String

AprobarProductoUnicaLlave = True

' FI2026 - era .MaxRows > 1 (mas de una fila de datos) -> .Rows > 2
If GrillaProductos.Rows > 2 Then
  
  NRow1% = Fila
  
  GrillaProductos.Col = 2
  FechaFila1 = GrillaProductos.Text
  
  GrillaProductos.Col = 4
  ProdFila1 = GrillaProductos.Text
  
  GrillaProductos.Col = 6
  LlaveFila1 = Trim$(Mid$(GrillaProductos.Text, 10, 30))
   
  For NRow2% = 1 To GrillaProductos.Rows - 1 ' FI2026 - era .MaxRows
  
    If NRow2% <> NRow1% Then
    
      GrillaProductos.Row = NRow2%
      GrillaProductos.Col = 2
      FechaFila2 = GrillaProductos.Text
      
      GrillaProductos.Col = 4
      ProdFila2 = GrillaProductos.Text
      
      GrillaProductos.Col = 6
      LlaveFila2 = Trim$(Mid$(GrillaProductos.Text, 10, 30))
    
      If ProdFila1 = ProdFila2 And LlaveFila1 = LlaveFila2 Then
      
        If IsDate(FechaFila1) And IsDate(FechaFila2) Then
          If CDate(FechaFila1) < CDate(FechaFila2) Then
            resprd = MsgBox("Existe un producto más reciente y con iguales características." + Chr(10) + "¿Desea continuar de todas formas?", 48 + 4, "Atención")
            If resprd <> vbYes Then
               AprobarProductoUnicaLlave = False
            End If
            Exit For
          End If
          
        Else
           resprd = MsgBox("Existe otro producto del cliente con iguales características, pero la fecha no es identificable." + Chr(10) + "¿Desea continuar de todas formas?", 48 + 4, "Atención")
           If resprd <> vbYes Then
              AprobarProductoUnicaLlave = False
           End If
           Exit For
         End If
         
      End If
      
    End If
  Next NRow2%
  
End If

End Function
