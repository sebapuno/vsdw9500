VERSION 4.00
Begin VB.Form Frm_ERVL 
   BackColor       =   &H00C8D0D4&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errores de Validación de Datos"
   ClientHeight    =   4740
   ClientLeft      =   1935
   ClientTop       =   2025
   ClientWidth     =   7200
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   400
      size            =   9.75
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Height          =   5145
   Left            =   1875
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Top             =   1680
   Width           =   7320
   Begin Threed.SSPanel Lb_Zona 
      Height          =   690
      Left            =   5850
      TabIndex        =   2
      Top             =   150
      Width           =   1215
      _version        =   65536
      _extentx        =   2143
      _extenty        =   1217
      _stockprops     =   15
      caption         =   "Regional"
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
   End

   Begin Threed.SSPanel Lb_Prod 
      Height          =   690
      Left            =   1725
      TabIndex        =   3
      Top             =   150
      Width           =   3915
      _version        =   65536
      _extentx        =   6906
      _extenty        =   1217
      _stockprops     =   15
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Begin VB.Label lb_NumPrd 
         Alignment       =   2  'Center
         Caption         =   "lb_NumPrd"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   375
         TabIndex        =   8
         Top             =   375
         Width           =   3165
      End

      Begin VB.Label lb_NomPrd 
         Alignment       =   2  'Center
         Caption         =   "lb_NomPrd"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   7
         Top             =   75
         Width           =   3765
      End
   End

   Begin Threed.SSPanel lb_Accion 
      Height          =   690
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   1365
      _version        =   65536
      _extentx        =   2408
      _extenty        =   1217
      _stockprops     =   15
      caption         =   "Modificación"
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
   End

   Begin Threed.SSPanel SSPanel1 
      Height          =   615
      Left            =   150
      TabIndex        =   6
      Top             =   3975
      Width           =   4665
      _version        =   65536
      _extentx        =   8229
      _extenty        =   1085
      _stockprops     =   15
      caption         =   $"ERVLDCTD.frx":0000
      forecolor       =   192
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Microsoft Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
   End

   Begin MSFlexGridLib.MSFlexGrid GridErr 
      Height          =   2865
      Left            =   150
      TabIndex        =   5
      Top             =   975
      Width           =   6915
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      ScrollBars      =   2
      HighLight       =   0   'False
   End

   Begin Threed.SSCommand Cancelar 
      Height          =   645
      Left            =   6225
      TabIndex        =   1
      Top             =   3975
      Width           =   840
      _version        =   65536
      _extentx        =   1482
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      font3d          =   3
      picture         =   "ERVLDCTD.frx":0074
   End

   Begin Threed.SSCommand Imprimir 
      Height          =   645
      Left            =   5100
      TabIndex        =   0
      Top             =   3975
      Width           =   870
      _version        =   65536
      _extentx        =   1535
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Imprimir"
      forecolor       =   8388608
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      font3d          =   3
      autosize        =   1
      picture         =   "ERVLDCTD.frx":0886
   End
End
Attribute VB_Name = "Frm_ERVL"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Cancelar_Click()

Unload Me

End Sub


Private Sub Form_Load()

Dim Success As Integer
Dim Data As String
Dim Vista As String
Dim Llave As String
Dim InicioData As Integer
Dim InicioLlave As Integer
Dim LargoData As Integer
Dim Correlativo As String
Dim LargoLlave As Integer
Dim PunteroVsdPrd As Recordset
Dim Estado As Boolean
Dim Producto As String

Me.Visible = False

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

PreparaGrid_Errores

If Not App.PrevInstance Then

  Rut = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  Correlativo = SacaValorINI("Parametros", "Param3", DIRECTORIO_01 + ARCHIVO_INI01)

  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set BaseSGT = OpenDatabase(DIRECTORIO_02 + ARCHIVO_MDB04)
  Set BaseVLDCTD = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB03)
    
  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  ' FI2026 - era: PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
  ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
  ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
  ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
  PunteroVsdPrd.Index = BaseVisado.TableDefs("Vsd_Prd").Indexes(0).Name
  PunteroVsdPrd.Seek "=", Rut, Producto, Val(Correlativo)
  
  Call ObtenerDatosDeCorte(Producto, InicioLlave, LargoLlave, InicioData, LargoData, NombreProducto, Vista)
  Data = Mid$(PunteroVsdPrd!Prd_Data, InicioData, LargoData)
  Llave = Mid$(PunteroVsdPrd!Prd_Data, InicioLlave, LargoLlave)
  Oficina = Mid$(PunteroVsdPrd!Prd_Data, 1, 3)
  
  lb_NomPrd.Caption = NombreProducto
  Operacion = Mid$(PunteroVsdPrd!Prd_Data, 40, 1)
  
  Select Case Operacion
     Case "O"
        CodOper = OPER_APERTURA   ' 0
        lb_Accion.Caption = "Apertura"
     Case "C"
        CodOper = OPER_CIERRE   ' 1
        lb_Accion.Caption = "Cierre"
     Case "R"
        CodOper = OPER_RENOVACION   ' 2
        lb_Accion.Caption = "Renovación"
     Case "M"
        CodOper = OPER_MODIFICACION   ' 3
        lb_Accion.Caption = "Modificación"
  End Select
  
   
  CtdLcd = Llave
  
  Data = Llave + Data
    
  Estado = False
   
  If Vista = "3030" Then
    MetReg = "Metro"
    Estado = Valida_3030(Data)
  ElseIf Vista = "3031" Then
    MetReg = "Metro"
    Estado = Valida_3031(Data)
  ElseIf Vista = "6391" Then   ' Regional
    MetReg = "Regional"
    Estado = Valida_6391(Data)
  Else
    Estado = True
  End If
  
  PunteroVsdPrd.Close
  
  BaseSGT.Close
  BaseVLDCTD.Close
  BaseVisado.Close
  
  Lb_Zona.Caption = MetReg

  CtdLcd = IIf(Len(Trim$(CtdLcd)) = 12 And Mid$(CtdLcd, 1) = "0", Mid$(CtdLcd, 2), CtdLcd)

  If Origen = "BCH" Then
    lb_NumPrd.Caption = Format$(CtdLcd, "0-000-00000-00")
  ElseIf Origen = "BAE" Then
    lb_NumPrd.Caption = Format$(CtdLcd, "0-00-00-000000")
  Else
    lb_NumPrd.Caption = CtdLcd
  End If
  
  If Estado Then
    GridErr.Col = 0
    GridErr.Row = 1
    If GridErr.Rows > 2 Or GridErr.Text <> "" Then
      Success = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZENOMOVE)
      Me.Visible = True
    
    Else
      '* No existen errores
      If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
         MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
                   "Error de Escritura"
      End If

      End
    
    End If
    
  End If
  
End If

End Sub


Public Sub PreparaGrid_Errores()

GridErr.Row = 0

GridErr.Col = 0
GridErr.Text = " Descripción Dato No Válido"
GridErr.FixedAlignment(0) = 0
GridErr.Col = 1
GridErr.Text = "Valor Incorrecto"
GridErr.FixedAlignment(1) = 2
GridErr.Col = 2
GridErr.Text = "Posición"
GridErr.FixedAlignment(2) = 2

GridErr.ColWidth(0) = 4095
GridErr.ColWidth(1) = 1560
GridErr.ColWidth(2) = 900
GridErr.ColWidth(3) = 1

GridErr.ColAlignment(1) = 2  ' 0 izq , 1 der , 2 Cen
GridErr.ColAlignment(2) = 2

End Sub




Private Sub GridErr_Click()

'If GridErr.Rows > 1 And GridErr.Row <> 0 Then
  'GridErr.Col = GridErr.Cols - 1
  'Ilumina_Row GridErr
'End If

End Sub

Private Sub Imprimir_Click()

Dim PunteroVld As Recordset

Screen.MousePointer = 11

'On Error GoTo ErrorDeImpresion
  
  Printer.FontName = "Courier New"
  Printer.FontSize = 12
  Printer.Print "BANCO DE CHILE"; Tab(54); "Fecha Emisión : " + Format$(Date, "dd/mmm/yyyy")
  Printer.Print Tab(54); "Hora Emisión  : " + Format$(Time, "hh:mm:ss")
  Printer.Print
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.Print Tab(25); "DETALLE ERRORES DE VALIDACIÓN"
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.Print
  Printer.Print
  Printer.FontSize = 10
  Printer.Print Tab(10); "Oficina"; Tab(22); ": " + Oficina
  Printer.Print Tab(10); "Rut Cliente"; Tab(22); ": " + Format$(Mid$(Rut, 1, Len(Rut) - 1), "###,###,###") + "-" + Right(Rut, 1)
  Printer.Print Tab(10); "Producto"; Tab(22); ": " + lb_NomPrd.Caption + "   N° " + lb_NumPrd.Caption
  Printer.Print Tab(10); "Acción"; Tab(22); ": " + lb_Accion.Caption
  Printer.Print Tab(10); "Origen"; Tab(22); ": " + Lb_Zona.Caption
  Printer.Print
  Printer.Print
  Printer.Print Tab(10); UCase("Descripción Dato No Válido" + Space(10) + "Valor Incorrecto" + Space(5) + "Posición")
  Printer.Print Tab(10); String(65, "-")
  
  For s% = 1 To GridErr.Rows - 1
    GridErr.Row = s%
    
    GridErr.Col = 0
    Printer.Print Tab(10); GridErr.Text;
    GridErr.Col = 1
    Printer.Print Tab(46); Left$(GridErr.Text + Space(10), 10);
    GridErr.Col = 2
    Printer.Print Tab(70); Trim$(GridErr.Text)
    
  Next s%
  
  '*** imprimir valores aceptables
  Set BaseVLDCTD = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB03)
  
  Printer.Print
  Printer.Print
  Printer.Print Tab(10); String(23, "-") + " VALORES CORRECTOS " + String(23, "-")
  
  For s% = 1 To GridErr.Rows - 1
    GridErr.Row = s%
    
    GridErr.Col = 3
    If GridErr.Text = 1 Then
    
      Printer.Print
      Printer.FontUnderline = True
      GridErr.Col = 0
      Printer.Print Tab(10); UCase(GridErr.Text);
      Printer.FontUnderline = False
    
      Tabla = IIf(UCase(MetReg) = "METRO", "tbl_ValMET", "tbl_ValREG")
    
      GridErr.Col = 2
      PosVta = Trim(Str(Val(GridErr.Text)))
          
      Set PunteroVld = BaseVLDCTD.OpenRecordset("Select * From " + Tabla + " Where CTD_Prod=" + TipoProd + " And CTD_Oper=" + CodOper + " And CTD_PosVista=" + PosVta + " Order by CTD_Argum Asc", dbOpenDynaset)
      '  CTD_Orig,CTD_Categ,CTD_PosVista
      If PunteroVld.RecordCount > 0 Then
        PunteroVld.MoveLast
        PunteroVld.MoveFirst
   
        For j% = 1 To PunteroVld.RecordCount
          Printer.Print Tab(10); PunteroVld!Ctd_Argum + " - " + Trim$(PunteroVld!Ctd_Descrip)
          PunteroVld.MoveNext
        Next j%
    
      End If
      
    End If
    
  Next s%
  Printer.Print
  Printer.Print
  Printer.Print
  Printer.FontSize = 12
  Printer.FontBold = True
  Printer.FontItalic = True
  Printer.Print Tab(15); "ANALIZAR ESTA SITUACIÓN CON EL EJECUTIVO ASOCIADO,"
  Printer.Print Tab(15); "  Y SI CORRESPONDE REPORTAR A LA MESA DE AYUDA."
  Printer.FontItalic = False
  Printer.FontBold = False
  Printer.FontSize = 12
  
  BaseVLDCTD.Close
  
  '**********
  
  Printer.EndDoc

On Error GoTo 0
  GoTo Fin
  
ErrorDeImpresion:
   i% = MsgBox("Error de Impresión", 5 + 48 + 0, "Impresión Cancelada")
   If i% = 4 Then
     Resume 0
   Else
     Printer.KillDoc
On Error GoTo 0
   End If
   
Fin:
  Screen.MousePointer = 0
  Cancelar.SetFocus


End Sub








Public Sub PintaDatos()


End Sub
