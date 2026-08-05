VERSION 4.00
Begin VB.Form ImprimeVisadosNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión Productos Ya Visados  "
   ClientHeight    =   6690
   ClientLeft      =   570
   ClientTop       =   1335
   ClientWidth     =   10890
   ControlBox      =   0   'False
   Height          =   7095
   Left            =   510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   Top             =   990
   Width           =   11010
   Begin Threed.SSFrame SSFrame1 
      Height          =   6645
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   10770
      _version        =   65536
      _extentx        =   18997
      _extenty        =   11721
      _stockprops     =   14
      caption         =   "Selección Productos a Imprimir ..."
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
      Begin VBX.SpreadSheet GrillaProductosVisados 
         DisplayRowHeaders=   0   'False
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         Height          =   5080
         InterfaceDesigner=   "IMPR_NEW.frx":0000
         Left            =   120
         MaxCols         =   8
         MaxRows         =   1
         OperationMode   =   1  'ReadOnly
         RetainSelBlock  =   0   'False
         ScrollBars      =   2  'Vertical
         ShadowColor     =   &H00C8D0D4&
         TabIndex        =   0
         Top             =   375
         Width           =   10470
      End
      Begin Threed.SSFrame Frame3D1 
         Height          =   960
         Left            =   150
         TabIndex        =   7
         Top             =   5550
         Width           =   4005
         _version        =   65536
         _extentx        =   7064
         _extenty        =   1693
         _stockprops     =   14
         caption         =   "Selección a Desplegar"
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
         Begin VBX.IntTextBox CantidadDias 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2640
            Max             =   365
            Min             =   1
            ReadOnly        =   0   'False
            SelStart        =   0
            Spin            =   -1  'True
            SpinInc         =   1
            SpinWrap        =   -1  'True
            TabIndex        =   3
            Text            =   "3"
            ThreeD          =   -1  'True
            ThreeDHighlightColor=   &H00FFFFFF&
            ThreeDIn        =   0   'False
            ThreeDShadowColor=   &H00808080&
            ThreeDShadowSize=   1
            Top             =   540
            Width           =   600
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
            caption         =   "Productos Visados &Hasta              Días"
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
            caption         =   "Productos Visados en el &Día"
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
      Begin Threed.SSCommand btn_Copiar 
         Height          =   645
         Left            =   6225
         TabIndex        =   9
         Top             =   5775
         Width           =   930
         _version        =   65536
         _extentx        =   1640
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Copiar"
         forecolor       =   32768
         font3d          =   1
         autosize        =   1
         picture         =   "IMPR_NEW.frx":0453
      End
      Begin Threed.SSCommand Cancelar 
         Height          =   645
         Left            =   8475
         TabIndex        =   5
         Top             =   5775
         Width           =   930
         _version        =   65536
         _extentx        =   1640
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   1
         picture         =   "IMPR_NEW.frx":076D
      End
      Begin Threed.SSCommand Imprimir 
         Height          =   645
         Left            =   7350
         TabIndex        =   4
         Top             =   5775
         Width           =   930
         _version        =   65536
         _extentx        =   1640
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Imprimir"
         forecolor       =   8388608
         font3d          =   1
         autosize        =   1
         picture         =   "IMPR_NEW.frx":0A87
      End
   End
   Begin VB.Label lb_OrdenGrilla 
      Caption         =   "lb_OrdenGrilla"
      Height          =   240
      Left            =   -75
      TabIndex        =   8
      Top             =   6525
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "ImprimeVisadosNew"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub btn_Copiar_Click()

Dim Ifil#, Icol#
Dim LineaSpread As String

Screen.MousePointer = 11

On Error GoTo ErrorArchivo

If GrillaProductosVisados.MaxRows > 0 Then

  NomArchivoHst = Format$(Now, "yyyymmdd") + ".XLS"
  
  Open DIRECTORIO_01 + NomArchivoHst For Output As #1
  
  For Ifil# = 0 To GrillaProductosVisados.MaxRows
    GrillaProductosVisados.Row = Ifil#
    LineaSpread = ""
    For Icol# = 1 To GrillaProductosVisados.MaxCols
      GrillaProductosVisados.Col = Icol#
      If Icol# = 1 Then
        LineaSpread = GrillaProductosVisados.Text
      Else
        LineaSpread = LineaSpread + Chr(9) + GrillaProductosVisados.Text
      End If
    Next Icol#
    
    Print #1, LineaSpread

  Next Ifil#
  
  Close #1
  
  MsgBox "El archivo se ha grabado como C:\DATA\VISADO\" + Format$(Now, "yyyymmdd") + ".XLS", 64, "Respaldo Productos Visados"
    
  '****

  GrillaProductosVisados.Row = 0
  GrillaProductosVisados.Row2 = GrillaProductosVisados.MaxRows
  GrillaProductosVisados.Col = 1
  GrillaProductosVisados.Col2 = GrillaProductosVisados.MaxCols
  GrillaProductosVisados.Action = 2

  GrillaProductosVisados.Action = 22

  GrillaProductosVisados.Action = 0
End If
  
Screen.MousePointer = 0
  
Exit Sub

ErrorArchivo:
   MsgBox "Para poder grabar debe cerrar el archivo abierto.", 16, "No se puede abrir el archivo"

  Screen.MousePointer = 0

End Sub


Private Sub Cancelar_Click()
  Unload Me
  'GrillaProductosVisados.OperationMode = 0
   
End Sub
Private Sub LlenarGrillaProductoVisado()
Dim Temporal As String
Dim FechaActual As String * 8
Dim PunteroVsdHst As Snapshot
Dim PunteroTabPrd As Recordset
Dim PunteroTabOper As Recordset
Dim FechaVsd As String
Dim FechPrd As String

  GrillaProductosVisados.MaxRows = 0
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  Set PunteroTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  PunteroTabOper.Index = PunteroTabOper.Indexes(0).Name
  If OpcionUnDia.Value Then
    FechaActual = Format$(Now, "yyyymmdd")
    Set PunteroVsdHst = BaseHistory.CreateSnapshot("SELECT * FROM tbl_Hst WHERE Hst_Fech='" + _
                                                   FechaActual + "';")
    'Set PunteroVsdHst = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Hst WHERE Hst_Fech='" + _
                                                                                         FechaActual + "';")
  Else
    FechaActual = Format$(Year(Date - CantidadDias.Val), "0000") + _
                             Format$(Month(Date - CantidadDias.Val), "00") + _
                             Format$(Day(Date - CantidadDias.Val), "00")
    Set PunteroVsdHst = BaseHistory.CreateSnapshot("SELECT * FROM tbl_Hst WHERE Val(Hst_Fech)>=" + _
                                                   FechaActual + ";")
    'Set PunteroVsdHst = BaseVisado.CreateSnapshot("SELECT * FROM Vsd_Hst WHERE Val(Hst_Fech)>=" + _
                                                                                         FechaActual + ";")
  End If
  Do While Not PunteroVsdHst.EOF
    GrillaProductosVisados.MaxRows = GrillaProductosVisados.MaxRows + 1
    GrillaProductosVisados.Row = GrillaProductosVisados.MaxRows
    PunteroTabPrd.Seek "=", PunteroVsdHst("Hst_Tipo")
    
    GrillaProductosVisados.Col = 1
    GrillaProductosVisados.TypeHAlign = 2
    GrillaProductosVisados.Text = Mid$(PunteroVsdHst("Hst_Data"), 1, 3)
    
    GrillaProductosVisados.Col = 2
    GrillaProductosVisados.TypeHAlign = 2
    If PunteroTabPrd("Prd_Pos_Eje") <> "0,0" Then
       GrillaProductosVisados.Text = Mid$(PunteroVsdHst("Hst_Data"), _
                                          Interpreta(PunteroTabPrd("Prd_Pos_Eje"), SIMBOLOPOSICION), _
                                          Interpreta(PunteroTabPrd("Prd_Pos_Eje"), SIMBOLOLARGO))
    Else
       GrillaProductosVisados.Text = "000"
    End If
    
    '* Fecha Producto *****
    GrillaProductosVisados.Col = 3
    GrillaProductosVisados.TypeHAlign = 2
    
    FechPrd = Trim$(ObtenerDescripcionGlosa(PunteroVsdHst("Hst_Tipo"), _
                                                                  Mid$(PunteroVsdHst("Hst_Data"), _
                                                                  Interpreta(PunteroTabPrd("Prd_Pos_Fec"), SIMBOLOPOSICION), _
                                                                  Interpreta(PunteroTabPrd("Prd_Pos_Fec"), SIMBOLOLARGO)), _
                                                                  PunteroTabPrd("Prd_Pic_Fec")))
    If InStr(FechPrd, "/") > 0 Then
      FechPrd = Mid$(FechPrd, 7, 4) + "/" + Mid$(FechPrd, 4, 2) + "/" + Mid$(FechPrd, 1, 2)
    End If
    GrillaProductosVisados.Text = Left(FechPrd + Space(10), 10)
    
    '* Nombre Producto *****
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
    GrillaProductosVisados.TypeHAlign = 2
    GrillaProductosVisados.Text = FormateaRut(PunteroVsdHst("Hst_Rut"))
    
    GrillaProductosVisados.Col = 7
    GrillaProductosVisados.TypeHAlign = 2
    If PunteroTabPrd("Prd_Pos_Dat") <> "0,0" Then
       GrillaProductosVisados.Text = ObtenerDescripcionGlosa(PunteroTabPrd("Prd_Producto"), _
                                                       Mid$(PunteroVsdHst("Hst_Data"), _
                                                       Interpreta(PunteroTabPrd("Prd_Pos_Dat"), SIMBOLOPOSICION), _
                                                       Interpreta(PunteroTabPrd("Prd_Pos_Dat"), SIMBOLOLARGO)), _
                                                       PunteroTabPrd("Prd_Pic_Dat"))
    Else
       GrillaProductosVisados.Text = ""
    End If
    
    GrillaProductosVisados.Col = 8
    GrillaProductosVisados.TypeHAlign = 2
    FechaVsd = Trim$(PunteroVsdHst("Hst_Fech"))
    GrillaProductosVisados.Text = Mid$(FechaVsd, 1, 4) + "/" + Mid$(FechaVsd, 5, 2) + "/" + Mid$(FechaVsd, 7, 2)

    
    PunteroVsdHst.MoveNext
  Loop
  PunteroTabPrd.Close
  PunteroVsdHst.Close
  
  lb_OrdenGrilla.Caption = ""
  
  If GrillaProductosVisados.MaxRows > 0 Then
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
  
  If KeyAscii = 13 Then
     OpcionTresDias.Value = True
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
     MsgBox "No se encontró la base de datos Hist_vsd.mdb", 16, "Archivo no encontrado"
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
  
  GrillaProductosVisados.MaxCols = 8
  GrillaProductosVisados.ColWidth(0) = 0
  GrillaProductosVisados.ColWidth(1) = 6
  GrillaProductosVisados.ColWidth(2) = 7
  GrillaProductosVisados.ColWidth(3) = 9
  GrillaProductosVisados.ColWidth(4) = 22
  GrillaProductosVisados.ColWidth(5) = 11
  GrillaProductosVisados.ColWidth(6) = 9
  GrillaProductosVisados.ColWidth(7) = 11
  GrillaProductosVisados.ColWidth(8) = 9
  
  GrillaProductosVisados.UserResizeCol = 2

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

Private Sub GrillaProductosVisados_DblClick(Col As Long, Row As Long)

Call OrdenarGrilla(1, Col, GrillaProductosVisados, 1)
  
GrillaProductosVisados.Row = 0
GrillaProductosVisados.Col = Col
lb_OrdenGrilla.Caption = GrillaProductosVisados.Text
  

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

On Error GoTo ErrorImpresora

  Screen.MousePointer = 11
  Printer.FontName = "Arial"
  Printer.FontName = "Courier New"
  Printer.FontSize = "9" '"10"

  If GrillaProductosVisados.MaxRows Mod 70 = 0 Then
     TotPag = GrillaProductosVisados.MaxRows \ 70
  Else
     TotPag = (GrillaProductosVisados.MaxRows \ 70) + 1
  End If
  FilaGrid = 0
  PagActual = 1
  
  Printer.Print "BANCO DE CHILE"; Tab(77); "Fecha : " + Format(Now, "dd/mm/yyyy");
  Printer.Print Tab(100); " 1/" + Trim(Str(TotPag))
  Printer.FontUnderline = True
  If OpcionUnDia.Value Then
    Printer.Print Tab(25); "INFORME DE PRODUCTOS VISADOS";
  Else
    Printer.Print Tab(25); "INFORME DE PRODUCTOS VISADOS HASTA " + CantidadDias.Text + " DIAS";
  End If
  Printer.FontUnderline = False
  Printer.Print Tab(77); "Hora  : "; Format$(Time, "hh:mm:ss")
  Printer.Print
  'Printer.Print "Oficina : "; IdOficina.Caption
  If Trim$(lb_OrdenGrilla.Caption) <> "" Then
     Printer.Print "Ordenado por : "; lb_OrdenGrilla.Caption
  Else
     Printer.Print
  End If
  
  Printer.Print
  Printer.Print String(103, "-")
  Printer.Print "OFI EJE FECHA      PRODUCTO                     ACCION         RUT          IDENTIFICACION   FCH-VISADO"
  Printer.Print String(103, "-")
  For I = 1 To GrillaProductosVisados.MaxRows
    FilaGrid = FilaGrid + 1
    GrillaProductosVisados.Row = I
    GrillaProductosVisados.Col = 1
    Cadena = GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 2
    Cadena = Cadena + GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 3
    Cadena = Cadena + GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 4
    Cadena = Cadena + Mid$(GrillaProductosVisados.Text, 1, 28) + _
                     String(28 - Len(Mid(GrillaProductosVisados.Text, 1, 28)), " ") + " "
    GrillaProductosVisados.Col = 5
    Cadena = Cadena + Mid$(GrillaProductosVisados.Text, 1, 14) + _
                     String(14 - Len(Mid(GrillaProductosVisados.Text, 1, 14)), " ") + " "
    GrillaProductosVisados.Col = 6
    Cadena = Cadena + GrillaProductosVisados.Text + " "
    GrillaProductosVisados.Col = 7
    Cadena = Cadena + Mid$(GrillaProductosVisados.Text, 1, 16) + _
                     String(16 - Len(Mid(GrillaProductosVisados.Text, 1, 16)), " ") + " "
    GrillaProductosVisados.Col = 8
    Cadena = Cadena + GrillaProductosVisados.Text
    
    Printer.Print Cadena
    
    If FilaGrid = 70 And GrillaProductosVisados.Row < GrillaProductosVisados.MaxRows Then
        Printer.EndDoc
        Printer.FontName = "Arial"
        Printer.FontName = "Courier New"
        Printer.FontSize = "9" '"10"
        PagActual = PagActual + 1
        FilaGrid = 0
        Printer.Print "BANCO DE CHILE"; Tab(77); "Fecha : "; Format(Now, "dd/mm/yyyy");
        Printer.Print Tab(100); Trim$(Str(PagActual)) + "/" + Trim(Str(TotPag))
        Printer.Print Tab(77); "Hora  : "; Format$(Time, "hh:mm:ss")
        Printer.Print
        Printer.Print
        Printer.Print String(103, "-")
        Printer.Print "OFI EJE FECHA      PRODUCTO                     ACCION         RUT          IDENTIFICACION   FCH-VISADO"
        Printer.Print String(103, "-")

    End If
  Next I
  Printer.EndDoc
  Screen.MousePointer = 0
  
Exit Sub

ErrorImpresora:
   If Err.Number = 482 Or Err.Number = 483 Or Err.Number = 484 Then
     valimp# = MsgBox("Problema para conectarse a la impresora." + Chr(10) + "La impresora no está conectada o no tiene papel.", 48 + 5, "Impresora Predeterminada")
     If valimp# = vbRetry Then
        Resume 0
     Else
        Screen.MousePointer = 0
     End If
   Else
     MsgBox "Error : " + Err.Number + Chr(10) + "Descripción : " + Err.Description
     Screen.MousePointer = 0
   End If
   
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
  If GrillaProductosVisados.MaxRows > 0 Then
    GrillaProductosVisados.SetFocus
  End If
End Sub


Private Sub InHabilitarDias(ByVal Estado As Boolean)
  CantidadDias.ReadOnly = Estado
  CantidadDias.Spin = False
  CantidadDias.Spin = True
End Sub


