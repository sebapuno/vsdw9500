VERSION 4.00
Begin VB.Form ImprimeProductoVisado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ImpresiÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n Productos Ya Visados"
   ClientHeight    =   4320
   ClientLeft      =   1440
   ClientTop       =   2160
   ClientWidth     =   9405
   ControlBox      =   0   'False
   Height          =   4725
   Left            =   1380
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Top             =   1815
   Width           =   9525
   Begin Threed.SSFrame SSFrame1 
      Height          =   4245
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   9330
      _version        =   65536
      _extentx        =   16457
      _extenty        =   7488
      _stockprops     =   14
      caption         =   "SelecciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n Productos a Imprimir ..."
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
      ' FI2026 - VBX.SpreadSheet migrado a MSFlexGrid (MSFLXGRD.OCX)
      Begin MSFlexGridLib.MSFlexGrid GrillaProductosVisados 
         Height          =   2385
         Left            =   120
         TabIndex        =   0
         Top             =   675
         Width           =   9110
         Cols            =   7
         Rows            =   1
         ScrollBars      =   2  'flexScrollBarVertical
         SelectionMode   =   1  'flexSelectionByRow
      End

      Begin Threed.SSFrame Frame3D1 
         Height          =   960
         Left            =   135
         TabIndex        =   7
         Top             =   3150
         Width           =   4005
         _version        =   65536
         _extentx        =   7064
         _extenty        =   1693
         _stockprops     =   14
         caption         =   "SelecciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n a Desplegar"
         forecolor       =   4210688
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         font3d          =   3
         ' FI2026 - VBX.IntTextBox migrado a VB.TextBox nativo
         Begin VB.TextBox CantidadDias 
            Height          =   255
            Left            =   2640
            TabIndex        =   3
            Top             =   540
            Width           =   600
            Text            =   "3"
         End

         Begin Threed.SSOption OpcionTresDias 
            Height          =   330
            Left            =   135
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   495
            Width           =   3705
            _version        =   65536
            _extentx        =   6535
            _extenty        =   582
            _stockprops     =   78
            caption         =   "Productos Visados &Hasta              DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­as"
            forecolor       =   128
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            font3d          =   3
         End

         Begin Threed.SSOption OpcionUnDia 
            Height          =   330
            Left            =   135
            TabIndex        =   1
            Top             =   225
            Width           =   2760
            _version        =   65536
            _extentx        =   4868
            _extenty        =   582
            _stockprops     =   78
            caption         =   "Productos Visados en el &DÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­a"
            forecolor       =   128
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            value           =   -1  'True
            font3d          =   3
         End
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Oficina"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   450
         Width           =   765
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "IdentificaciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   6
         Left            =   7770
         TabIndex        =   13
         Top             =   450
         Width           =   1200
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Rut"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   6640
         TabIndex        =   12
         Top             =   450
         Width           =   1110
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "AcciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   5340
         TabIndex        =   11
         Top             =   450
         Width           =   1290
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Ejecutivo"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   900
         TabIndex        =   10
         Top             =   450
         Width           =   765
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Fecha"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         Top             =   450
         Width           =   960
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   2660
         TabIndex        =   8
         Top             =   450
         Width           =   2670
      End

      Begin Threed.SSCommand Cancelar 
         Height          =   645
         Left            =   8440
         TabIndex        =   5
         Top             =   3495
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   1
         picture         =   "IMPRPROD.frx":0004
      End

      Begin Threed.SSCommand Imprimir 
         Height          =   645
         Left            =   7630
         TabIndex        =   4
         Top             =   3495
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Imprimir"
         forecolor       =   8388608
         font3d          =   1
         autosize        =   1
         picture         =   "IMPRPROD.frx":0816
      End
   End

   Begin VB.Label lb_OrdenGrilla 
      Caption         =   "lb_OrdenGrilla"
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "ImprimeProductoVisado"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Cancelar_Click()
  Unload Me
  'GrillaProductosVisados.OperationMode = 0
   
End Sub
Private Sub LlenarGrillaProductoVisado()
Dim Temporal As String
Dim FechaActual As String * 8
Dim PunteroVsdHst As Recordset
Dim PunteroTabPrd As Recordset
Dim PunteroTabOper As Recordset
  ' FI2026 - MaxRows=0 (vaciar) -> Rows=1: MSFlexGrid no admite Rows=0 con
  ' FixedRows=1; 1 = solo la fila fija (oculta), sin filas de datos
  GrillaProductosVisados.Rows = 1
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  ' FI2026 - PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  If OpcionUnDia.Value Then
    FechaActual = Format$(Now, "yyyymmdd")
    Set PunteroVsdHst = BaseHistory.OpenRecordset("SELECT * FROM tbl_Hst WHERE Hst_Fech='" + _
                                                   FechaActual + "';", dbOpenSnapshot)
    'Set PunteroVsdHst = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Hst WHERE Hst_Fech='" + _
                                                                                         FechaActual + "';")
  Else
    FechaActual = Format$(Year(Date - Val(CantidadDias.Text)), "0000") + _
                             Format$(Month(Date - Val(CantidadDias.Text)), "00") + _
                             Format$(Day(Date - Val(CantidadDias.Text)), "00")
    Set PunteroVsdHst = BaseHistory.OpenRecordset("SELECT * FROM tbl_Hst WHERE Val(Hst_Fech)>=" + _
                                                   FechaActual + ";", dbOpenSnapshot)
    'Set PunteroVsdHst = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Hst WHERE Val(Hst_Fech)>=" + _
                                                                                         FechaActual + ";")
  End If
  Do While Not PunteroVsdHst.EOF
    GrillaProductosVisados.Rows = GrillaProductosVisados.Rows + 1
    GrillaProductosVisados.Row = GrillaProductosVisados.Rows - 1 ' FI2026 - ultima fila de datos (era .MaxRows)
    PunteroTabPrd.Seek "=", PunteroVsdHst("Hst_Tipo")
    
    GrillaProductosVisados.Col = 1
    ' FI2026 - GrillaProductosVisados.TypeHAlign = 2 migrado a GrillaProductosVisados.CellAlignment
    GrillaProductosVisados.CellAlignment = 4  ' flexAlignCenterCenter
    GrillaProductosVisados.Text = Mid$(PunteroVsdHst("Hst_Data"), 1, 3)
    
    GrillaProductosVisados.Col = 2
    ' FI2026 - GrillaProductosVisados.TypeHAlign = 2 migrado a GrillaProductosVisados.CellAlignment
    GrillaProductosVisados.CellAlignment = 4  ' flexAlignCenterCenter
    If PunteroTabPrd("Prd_Pos_Eje") <> "0,0" Then
       GrillaProductosVisados.Text = Mid$(PunteroVsdHst("Hst_Data"), _
                                          Interpreta(PunteroTabPrd("Prd_Pos_Eje"), SIMBOLOPOSICION), _
                                          Interpreta(PunteroTabPrd("Prd_Pos_Eje"), SIMBOLOLARGO))
    Else
       GrillaProductosVisados.Text = "000"
    End If
    GrillaProductosVisados.Col = 3
    ' FI2026 - GrillaProductosVisados.TypeHAlign = 2 migrado a GrillaProductosVisados.CellAlignment
    GrillaProductosVisados.CellAlignment = 4  ' flexAlignCenterCenter
    GrillaProductosVisados.Text = Trim$(ObtenerDescripcionGlosa(PunteroVsdHst("Hst_Tipo"), _
                                                                  Mid$(PunteroVsdHst("Hst_Data"), _
                                                                  Interpreta(PunteroTabPrd("Prd_Pos_Fec"), SIMBOLOPOSICION), _
                                                                  Interpreta(PunteroTabPrd("Prd_Pos_Fec"), SIMBOLOLARGO)), _
                                                                  PunteroTabPrd("Prd_Pic_Fec")))
    GrillaProductosVisados.Col = 4
    GrillaProductosVisados.Text = Trim$(PunteroTabPrd("Prd_Descrip"))
    If Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOPOSICION) = -1 Then
      Temporal = PunteroTabPrd("Prd_Pos_Act")
    Else
      Temporal = Mid$(PunteroVsdHst("Hst_Data"), _
                        Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOPOSICION), _
                        Interpreta(PunteroTabPrd("Prd_Pos_Act"), SIMBOLOLARGO))
    End If
    GrillaProductosVisados.Col = 5
    PunteroTabOper.Seek "=", PunteroVsdHst("Hst_Tipo"), Temporal
    If Not PunteroTabOper.NoMatch Then
      GrillaProductosVisados.Text = PunteroTabOper("Oper_Glosa")
    Else
      GrillaProductosVisados.Text = Temporal
    End If
    GrillaProductosVisados.Col = 6
    ' FI2026 - GrillaProductosVisados.TypeHAlign = 2 migrado a GrillaProductosVisados.CellAlignment
    GrillaProductosVisados.CellAlignment = 4  ' flexAlignCenterCenter
    GrillaProductosVisados.Text = FormateaRut(CStr(PunteroVsdHst("Hst_Rut")))
    
    GrillaProductosVisados.Col = 7
    ' FI2026 - GrillaProductosVisados.TypeHAlign = 2 migrado a GrillaProductosVisados.CellAlignment
    GrillaProductosVisados.CellAlignment = 4  ' flexAlignCenterCenter
    If PunteroTabPrd("Prd_Pos_Dat") <> "0,0" Then
       GrillaProductosVisados.Text = ObtenerDescripcionGlosa(PunteroTabPrd("Prd_Producto"), _
                                                       Mid$(PunteroVsdHst("Hst_Data"), _
                                                       Interpreta(PunteroTabPrd("Prd_Pos_Dat"), SIMBOLOPOSICION), _
                                                       Interpreta(PunteroTabPrd("Prd_Pos_Dat"), SIMBOLOLARGO)), _
                                                       PunteroTabPrd("Prd_Pic_Dat"))
    Else
       GrillaProductosVisados.Text = ""
    End If
    PunteroVsdHst.MoveNext
  Loop
  PunteroTabPrd.Close
  PunteroVsdHst.Close
  
  lb_OrdenGrilla.Caption = ""
  
  If GrillaProductosVisados.Rows > 1 Then ' FI2026 - 'hay datos': era .MaxRows > 0
    Imprimir.Enabled = True
    Call OrdenarGrilla(1, 3, GrillaProductosVisados, 1)
  Else
    Imprimir.Enabled = False
  End If
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub

Private Sub CantidadDias_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub

Private Sub CantidadDias_LostFocus()
  If (CantidadDias.Text < 1) Or (CantidadDias.Text > 365) Then
    CantidadDias.Text = 3
  End If
End Sub
Private Sub Form_Activate()
  Call DarFocoGrillaProductosVisados
  Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim FechaPasada As String * 8
  Screen.MousePointer = 11
  Principal.Enabled = False
  '---------
  If Dir("c:\data\visado\hist_vsd.mdb") <> "HIST_VSD.MDB" Then
     MsgBox "No se encontrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³ la base de datos Hist_vsd.mdb", 16, "Archivo no encontrado"
     Screen.MousePointer = 0
     Frame3D1.Enabled = False
     Exit Sub
  End If
  
  Set BaseHistory = OpenDatabase(DIRECTORIO_01 + "Hist_vsd.mdb")
  Llena_Historico
  FechaPasada = Format$(Year(Date - 365), "0000") + Format$(Month(Date - 365), "00") + _
                             Format$(Day(Date - 365), "00")
  Frame3D1.Enabled = True
   
  Call CentrarVentana(Me)
  
  'FechaPasada = Format$(Year(Date - 31), "0000") + Format$(Month(Date - 31), "00") + _
                             Format$(Day(Date - 31), "00")
  'BaseVisado.Execute "DELETE * FROM Vsd_Hst WHERE Val(Hst_Fech)<" + FechaPasada + ";"
  'Call CentrarVentana(Me)
  
  lb_OrdenGrilla.Caption = ""
  Call ColocaTitulos
  
  ' FI2026 - MaxCols=7 (Spread: columnas de DATOS 1..7) -> Cols=8
  GrillaProductosVisados.Cols = 8
  ' FI2026 - el Spread original tenia DisplayColHeaders=0 y DisplayRowHeaders=0:
  ' se ocultan la columna fija 0 y la fila fija 0 de MSFlexGrid
  GrillaProductosVisados.ColWidth(0) = 0
  GrillaProductosVisados.RowHeight(0) = 0
  GrillaProductosVisados.ColWidth(1) = 720  ' FI2026 - eran 6 caracteres (Spread)
  GrillaProductosVisados.ColWidth(2) = 840  ' FI2026 - eran 7 caracteres (Spread)
  GrillaProductosVisados.ColWidth(3) = 960  ' FI2026 - eran 8 caracteres (Spread)
  GrillaProductosVisados.ColWidth(4) = 2604  ' FI2026 - eran 21.7 caracteres (Spread)
  GrillaProductosVisados.ColWidth(5) = 1200  ' FI2026 - eran 10 caracteres (Spread)
  GrillaProductosVisados.ColWidth(6) = 1080  ' FI2026 - eran 9 caracteres (Spread)
  GrillaProductosVisados.ColWidth(7) = 1320  ' FI2026 - eran 11 caracteres (Spread)
  
  'IdOficina.Caption = GlosaOficina(Oficina)
  Call LlenarGrillaProductoVisado
  
  '*****
  BaseHistory.Execute "DELETE * FROM tbl_Hst WHERE Val(Hst_Fech)<" + FechaPasada + ";"
  BaseHistory.Close
  
  DBEngine.CompactDatabase "C:\data\visado\hist_vsd.mdb", "C:\data\visado\hist_tmp.mdb"
  If Dir("C:\data\visado\hist_tmp.mdb") = "HIST_TMP.MDB" Then
     Kill "C:\data\visado\hist_vsd.mdb"
     FileCopy "C:\data\visado\hist_tmp.mdb", "C:\data\visado\hist_vsd.mdb"
     Kill "C:\data\visado\hist_tmp.mdb"
     Kill "C:\data\visado\hist_tmp.ldb"
  End If
  
  Set BaseHistory = OpenDatabase(DIRECTORIO_01 + "Hist_vsd.mdb")
  '*****
  Screen.MousePointer = 0
   
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Principal.Enabled = True
  BaseHistory.Close

End Sub

Private Sub GrillaProductosVisados_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub

Private Sub Imprimir_Click()
Dim I As Integer
Dim Cadena As String
Dim TotPag As Integer
Dim PagActual As Integer
Dim FilaGrid As Integer

  Screen.MousePointer = 11
  Printer.FontName = "Courier New"
  Printer.FontSize = "10"
  ' FI2026 - (.Rows - 1) = cantidad de filas de datos (era .MaxRows)
  If (GrillaProductosVisados.Rows - 1) Mod 60 = 0 Then
     TotPag = (GrillaProductosVisados.Rows - 1) \ 60 ' FI2026 - era .MaxRows
  Else
     TotPag = ((GrillaProductosVisados.Rows - 1) \ 60) + 1 ' FI2026 - era .MaxRows
  End If
  FilaGrid = 0
  PagActual = 1
  
  Printer.Print "BANCO DE CHILE"; Tab(67); "Fecha : "; Format(Now, "dd/mm/yyyy");
  Printer.Print Tab(90); " 1/" + Trim(Str(TotPag))
  Printer.FontUnderline = True
  If OpcionUnDia.Value Then
    Printer.Print Spc(18); "INFORME DE PRODUCTOS VISADOS";
  Else
    Printer.Print Spc(18); "INFORME DE PRODUCTOS VISADOS HASTA " + CantidadDias.Text + " DIAS";
  End If
  Printer.FontUnderline = False
  Printer.Print Tab(67); "Hora  : "; Format$(Time, "hh:mm:ss")
  Printer.Print
  'Printer.Print "Oficina : "; IdOficina.Caption
  If Trim$(lb_OrdenGrilla.Caption) <> "" Then
     Printer.Print "Ordenado por : "; lb_OrdenGrilla.Caption
  Else
     Printer.Print
  End If
  
  Printer.Print
  Printer.Print String(95, "-")
  Printer.Print "OFI EJE FECHA      PRODUCTO                       ACCION          RUT          IDENTIFICACION"
  Printer.Print String(95, "-")
  For I = 1 To GrillaProductosVisados.Rows - 1 ' FI2026 - era .MaxRows
    FilaGrid = FilaGrid + 1
    GrillaProductosVisados.Row = I
    GrillaProductosVisados.Col = 1
    Cadena = GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 2
    Cadena = Cadena + GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 3
    Cadena = Cadena + GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 4
    Cadena = Cadena + Mid$(GrillaProductosVisados.Text, 1, 30) + _
                     String(30 - Len(Mid(GrillaProductosVisados.Text, 1, 30)), " ") + " "
    GrillaProductosVisados.Col = 5
    Cadena = Cadena + Mid$(GrillaProductosVisados.Text, 1, 14) + _
                     String(14 - Len(Mid(GrillaProductosVisados.Text, 1, 14)), " ") + " "
    GrillaProductosVisados.Col = 6
    Cadena = Cadena + GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 7
    Cadena = Cadena + Mid$(GrillaProductosVisados.Text, 1, 15)
    Printer.Print Cadena
    
    ' FI2026 - .Rows - 1: era .MaxRows (ultima fila de datos)
    If FilaGrid = 60 And GrillaProductosVisados.Row < GrillaProductosVisados.Rows - 1 Then
        Printer.EndDoc
        Printer.FontName = "Courier New"
        Printer.FontSize = "10"
        PagActual = PagActual + 1
        FilaGrid = 0
        Printer.Print "BANCO DE CHILE"; Tab(67); "Fecha : "; Format(Now, "dd/mm/yyyy");
        Printer.Print Tab(90); Trim$(Str(PagActual)) + "/" + Trim(Str(TotPag))
        Printer.Print Tab(67); "Hora  : "; Format$(Time, "hh:mm:ss")
        Printer.Print
        Printer.Print
        Printer.Print String(95, "-")
        Printer.Print "OFI EJE FECHA      PRODUCTO                       ACCION          RUT          IDENTIFICACION"
        Printer.Print String(95, "-")

    End If
  Next I
  Printer.EndDoc
  Screen.MousePointer = 0
End Sub
Private Sub Imprimir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub OpcionTresDias_Click(Value As Integer)
  Screen.MousePointer = 11
  Call InHabilitarDias(True)
  Call LlenarGrillaProductoVisado
  Call DarFocoGrillaProductosVisados
  Screen.MousePointer = 0
End Sub
Private Sub OpcionTresDias_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub OpcionUnDia_Click(Value As Integer)
  Screen.MousePointer = 11
  Call LlenarGrillaProductoVisado
  Call DarFocoGrillaProductosVisados
  Call InHabilitarDias(False)
  Screen.MousePointer = 0
End Sub
Private Sub OpcionUnDia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub DarFocoGrillaProductosVisados()
  If GrillaProductosVisados.Rows > 1 Then ' FI2026 - 'hay datos': era .MaxRows > 0
    GrillaProductosVisados.SetFocus
  End If
End Sub


Private Sub InHabilitarDias(ByVal Estado As Boolean)
  ' FI2026 - ReadOnly (VBX.IntTextBox) -> Locked (VB.TextBox nativo)
  CantidadDias.Locked = Estado
  ' FI2026_TODO - Spin (VBX.IntTextBox) sin equivalente en VB.TextBox nativo: el
  ' control original traia flechas de incremento/decremento. El par False/True
  ' era un refresco del control. Requiere un UpDown (MSCOMCT2.OCX) para replicarlo.
  'CantidadDias.Spin = False
  'CantidadDias.Spin = True
End Sub

Private Sub Titulo_Click(Index As Integer)
  Call OrdenarGrilla(1, Index + 1, GrillaProductosVisados, 1)
  
  Select Case Index
    Case 0
        lb_OrdenGrilla.Caption = "Oficina"
    Case 1
        lb_OrdenGrilla.Caption = "Ejecutivo"
    Case 2
        lb_OrdenGrilla.Caption = "Fecha"
    Case 3
        lb_OrdenGrilla.Caption = "Producto"
    Case 4
        lb_OrdenGrilla.Caption = "AcciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n"
    Case 5
        lb_OrdenGrilla.Caption = "Rut"
    Case 6
        lb_OrdenGrilla.Caption = "IdentificaciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n"
  End Select
 

End Sub

Public Sub ColocaTitulos()

Titulo(0).Left = 120
Titulo(0).Width = 740
Titulo(1).Left = 870
Titulo(1).Width = 840
Titulo(2).Left = 1720
Titulo(2).Width = 960
Titulo(3).Left = 2700
Titulo(3).Width = 2640
Titulo(4).Left = 5360
Titulo(4).Width = 1170
Titulo(5).Left = 6540
Titulo(5).Width = 1080
Titulo(6).Left = 7640
Titulo(6).Width = 1300
    
    
End Sub
