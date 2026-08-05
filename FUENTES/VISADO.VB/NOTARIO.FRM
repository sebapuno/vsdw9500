VERSION 4.00
Begin VB.Form ConfiguraHonorarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación Honorarios Notario"
   ClientHeight    =   3045
   ClientLeft      =   2880
   ClientTop       =   2415
   ClientWidth     =   6720
   ControlBox      =   0   'False
   Height          =   3450
   Left            =   2820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   Top             =   2070
   Width           =   6840
   Begin Threed.SSFrame SSFrame1 
      Height          =   3045
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   6645
      _version        =   65536
      _extentx        =   11721
      _extenty        =   5371
      _stockprops     =   14
      caption         =   "Selección de Notario ..."
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
      Begin VBX.SpreadSheet GrillaHonorarios 
         BorderStyle     =   0  'None
         DAutoSizeCols   =   0  'Off
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5  'Next
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         GridColor       =   &H00808080&
         Height          =   1440
         InterfaceDesigner=   "NOTARIO.frx":0000
         Left            =   120
         MaxCols         =   4
         MaxRows         =   7
         Protect         =   0   'False
         ScrollBars      =   2  'Vertical
         ShadowColor     =   &H00C8D0D4&
         ShadowText      =   &H00404000&
         TabIndex        =   1
         Top             =   840
         Width           =   6425
      End
      Begin VB.Label Instrucciones 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Borrar Actual Notario."
         BeginProperty Font 
            name            =   "Small Fonts"
            charset         =   1
            weight          =   400
            size            =   6.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   165
         Index           =   3
         Left            =   750
         TabIndex        =   13
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label Instrucciones 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": Ingresar Nuevo Notario."
         BeginProperty Font 
            name            =   "Small Fonts"
            charset         =   1
            weight          =   400
            size            =   6.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   165
         Index           =   2
         Left            =   750
         TabIndex        =   12
         Top             =   2340
         Width           =   1560
      End
      Begin VB.Label Instrucciones 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[SUPR]"
         BeginProperty Font 
            name            =   "Small Fonts"
            charset         =   1
            weight          =   400
            size            =   6.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   165
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Instrucciones 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[INSERT]"
         BeginProperty Font 
            name            =   "Small Fonts"
            charset         =   1
            weight          =   400
            size            =   6.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2340
         Width           =   600
      End
      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Costo Renovación"
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
         Height          =   510
         Index           =   3
         Left            =   4965
         TabIndex        =   7
         Top             =   315
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Costo Otorgamiento"
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
         Height          =   510
         Index           =   2
         Left            =   3630
         TabIndex        =   6
         Top             =   315
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Nombre"
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
         Left            =   1215
         TabIndex        =   5
         Top             =   450
         Width           =   2400
      End
      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Rut"
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
         Left            =   120
         TabIndex        =   4
         Top             =   450
         Width           =   1080
      End
      Begin Threed.SSCommand Cancelar 
         Height          =   615
         Left            =   5670
         TabIndex        =   3
         Top             =   2340
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   3
         picture         =   "NOTARIO.frx":0004
      End
      Begin Threed.SSCommand Aceptar 
         Height          =   615
         Left            =   4770
         TabIndex        =   2
         Top             =   2340
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Aceptar"
         forecolor       =   8421376
         font3d          =   3
         picture         =   "NOTARIO.frx":031E
      End
      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         Height          =   510
         Index           =   4
         Left            =   1215
         TabIndex        =   8
         Top             =   315
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Encabezado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
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
         Height          =   510
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   315
         Width           =   1080
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "ConfiguraHonorarios"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Editando As Boolean
Dim FilasVisibles As Integer
Dim BaseNotarios As Database
Dim CeldaModificada As Integer
Private Sub Aceptar_Click()
Dim I As Integer
Dim Criterio As String
Dim RutNotario As String
Dim PunteroNotarios As Recordset
  I = MsgBox("¿ Son los Datos Correctos ?", 4 + 32 + 256, "Verificar Modificación")
  If I = 6 Then
    Screen.MousePointer = 11
    For I = 1 To GrillaHonorarios.MaxRows
      GrillaHonorarios.Row = I
      If Not GrillaHonorarios.RowHidden Then
        If TodaLaFilaLlena Then
          GrillaHonorarios.Col = 1
          RutNotario = DesformateaRut(GrillaHonorarios.Text)
          GrillaHonorarios.Col = 5
          GrillaHonorarios.Text = Val(Mid(RutNotario, 1, Len(RutNotario) - 1))
        Else
          MsgBox "Notario Con Datos Incompletos", 0 + 16, "Faltan Datos"
          Call IrCeldaEspecifica(I, 1)
          GrillaHonorarios.SetFocus
          Screen.MousePointer = 0
          Exit Sub
        End If
      End If
    Next I
    Set PunteroNotarios = BaseNotarios.OpenRecordset("Hono_Rut", dbOpenDynaset)
    For I = 1 To GrillaHonorarios.MaxRows
      GrillaHonorarios.Row = I
      GrillaHonorarios.Col = 5
      Criterio = "Rut=" + Trim$(GrillaHonorarios.Text)
      PunteroNotarios.FindFirst Criterio
      If PunteroNotarios.NoMatch Then
          PunteroNotarios.AddNew
          Call IngresarDatosTabla(PunteroNotarios)
          PunteroNotarios.Update
      Else
        If GrillaHonorarios.RowHidden Then
          PunteroNotarios.Delete
        Else
          PunteroNotarios.Edit
          Call IngresarDatosTabla(PunteroNotarios)
          PunteroNotarios.Update
        End If
      End If
    Next I
    PunteroNotarios.Close
    Screen.MousePointer = 0
  End If
End Sub
Private Sub Aceptar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub Cancelar_Click()
  Unload Me
End Sub

Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub

Private Sub Encabezado_Click(Index As Integer)
  Select Case Index
    Case 0, 5
      Call OrdenarGrilla(1, 1, GrillaHonorarios, 1)
    Case 1, 4
      Call OrdenarGrilla(1, 2, GrillaHonorarios, 1)
    Case Else
      Call OrdenarGrilla(1, Index + 1, GrillaHonorarios, 1)
  End Select
End Sub
Private Sub Form_Activate()
  Editando = False
  FilasVisibles = GrillaHonorarios.MaxRows
  CeldaModificada = False
  Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
Dim PunteroNotarios As Recordset
  Screen.MousePointer = 11
  Principal.Enabled = False
  Call CentrarVentana(Me)
  GrillaHonorarios.MaxCols = 5
  GrillaHonorarios.MaxRows = 0
  GrillaHonorarios.ColWidth(1) = 9
  GrillaHonorarios.ColWidth(2) = 20
  GrillaHonorarios.ColWidth(3) = 11
  GrillaHonorarios.ColWidth(4) = 11
  GrillaHonorarios.ColWidth(5) = 0
  GrillaHonorarios.Col = 5
  GrillaHonorarios.ColHidden = True
  GrillaHonorarios.UserResizeRow = 2
  Set BaseNotarios = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB04)
  Set PunteroNotarios = BaseNotarios.OpenRecordset("Hono_Rut", dbOpenTable)
  While Not PunteroNotarios.EOF
    GrillaHonorarios.MaxRows = GrillaHonorarios.MaxRows + 1
    GrillaHonorarios.Row = GrillaHonorarios.MaxRows
    Call FormateaCeldasGrilla
    GrillaHonorarios.Col = 1
    GrillaHonorarios.Text = Format(PunteroNotarios("Rut"), "00,000,000") + "-" + PunteroNotarios("DV")
    GrillaHonorarios.Col = 2
    GrillaHonorarios.Text = StrConv(PunteroNotarios("Nombre"), 3)
    GrillaHonorarios.Col = 3
    GrillaHonorarios.Text = PunteroNotarios("Otorgam")
    GrillaHonorarios.Col = 4
    GrillaHonorarios.Text = PunteroNotarios("Novacion")
    GrillaHonorarios.Col = 5
    GrillaHonorarios.Text = PunteroNotarios("Rut")
    PunteroNotarios.MoveNext
  Wend
  PunteroNotarios.Close
  Call OrdenarGrilla(1, 2, GrillaHonorarios, 1)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  BaseNotarios.Close
  Principal.Enabled = True
End Sub

Private Sub GrillaHonorarios_EditMode(Col As Long, Row As Long, Mode As Integer, ChangeMade As Integer)
  If Mode = 1 Then
    Editando = True
  Else
    Editando = False
    CeldaModificada = ChangeMade
  End If
End Sub
Private Sub GrillaHonorarios_KeyDown(KeyCode As Integer, Shift As Integer)
Dim InsertarFila As Boolean
  If (Not Editando) And (Shift = 0) Then
    Select Case KeyCode
      Case 27
        Call Cancelar_Click
      Case 45
        InsertarFila = True
        If GrillaHonorarios.MaxRows > 0 Then
          GrillaHonorarios.Row = GrillaHonorarios.MaxRows
          If Not GrillaHonorarios.RowHidden Then
            If Not TodaLaFilaLlena Then
              InsertarFila = False
            End If
          End If
        End If
        If InsertarFila Then
          GrillaHonorarios.MaxRows = GrillaHonorarios.MaxRows + 1
          GrillaHonorarios.Row = GrillaHonorarios.MaxRows
          Call FormateaCeldasGrilla
          Call IrCeldaEspecifica(GrillaHonorarios.MaxRows, 1)
          FilasVisibles = FilasVisibles + 1
        End If
      Case 46
        If FilasVisibles > 0 Then
          GrillaHonorarios.Row = GrillaHonorarios.ActiveRow
          If TodaLaFilaVacia Then
            GrillaHonorarios.MaxRows = GrillaHonorarios.MaxRows - 1
            FilasVisibles = FilasVisibles - 1
          Else
            I% = MsgBox("¿ Seguro de Eliminar a Dicho Notario ?", 4 + 32 + 256, "Verificar Borrado")
            If I% = 6 Then
              GrillaHonorarios.RowHidden = True
              FilasVisibles = FilasVisibles - 1
            End If
          End If
        End If
    End Select
  End If
End Sub
Private Sub FormateaCeldasGrilla()
  GrillaHonorarios.Col = 1
  GrillaHonorarios.CellType = 4
  GrillaHonorarios.TypePicMask = "NN.NNN.NNN-N"
  GrillaHonorarios.TypeHAlign = 2
  GrillaHonorarios.Col = 3
  GrillaHonorarios.CellType = 3
  GrillaHonorarios.TypeHAlign = 2
  GrillaHonorarios.Col = 4
  GrillaHonorarios.CellType = 3
  GrillaHonorarios.TypeHAlign = 2
End Sub
Private Function TodaLaFilaVacia() As Boolean
  TodaLaFilaVacia = True
  For I% = 1 To 4
    GrillaHonorarios.Col = I%
    If GrillaHonorarios.Text <> "" Then
      TodaLaFilaVacia = False
      Exit For
    End If
  Next I%
End Function
Private Function TodaLaFilaLlena() As Boolean
  TodaLaFilaLlena = True
  For I% = 1 To 4
    GrillaHonorarios.Col = I%
    If GrillaHonorarios.Text = "" Then
      TodaLaFilaLlena = False
      Exit For
    End If
  Next I%
End Function
Private Sub GrillaHonorarios_LeaveCell(Col As Long, Row As Long, NewCol As Long, NewRow As Long, Cancel As Integer)
Dim RutNotario As String
Dim DigitoVerificador As String * 1
Dim NumeroTemporal As String
  GrillaHonorarios.Col = Col
  GrillaHonorarios.Row = Row
  If (GrillaHonorarios.Text <> "") And (CeldaModificada) Then
    CeldaModificada = False
    Select Case Col
      Case 1
        NumeroTemporal = Trim$(DesformateaRut(GrillaHonorarios.Text))
        RutNotario = Mid$(NumeroTemporal, 1, Len(NumeroTemporal) - 1)
        DigitoVerificador = UCase$(Mid(NumeroTemporal, Len(NumeroTemporal), 1))
        If (Val(RutNotario) <> 0) Or (DigitoVerificador <> "0") Then
          If ObtenerDigitoVerificador(RutNotario) = DigitoVerificador Then
            GrillaHonorarios.Text = FormateaRut(NumeroTemporal)
            Call IrCeldaEspecifica(NewRow, NewCol)
          Else
            MsgBox "Dígito Verificador Erroneo", 0 + 16, "Error de Validación"
            GrillaHonorarios.Text = ""
            Call IrCeldaEspecifica(Row, Col)
          End If
        Else
          MsgBox "No es Rut Válido", 0 + 16, "Faltan Datos"
          GrillaHonorarios.Text = ""
          Call IrCeldaEspecifica(Row, Col)
        End If
      Case 2
        GrillaHonorarios.Text = StrConv(GrillaHonorarios.Text, 3)
        Call IrCeldaEspecifica(NewRow, NewCol)
    End Select
  End If
End Sub
Private Sub IrCeldaEspecifica(ByVal Fila As Integer, ByVal Columna As Integer)
  GrillaHonorarios.Row = Fila
  GrillaHonorarios.Col = Columna
  GrillaHonorarios.Position = 5
  GrillaHonorarios.Action = 0
  GrillaHonorarios.Action = 1
End Sub

Private Sub IngresarDatosTabla(PunteroTabla As Recordset)
Dim RutTemporal As String
  GrillaHonorarios.Col = 1
  RutTemporal = Trim$(DesformateaRut(GrillaHonorarios.Text))
  PunteroTabla("DV") = UCase$(Mid(RutTemporal, Len(RutTemporal), 1))
  PunteroTabla("Rut") = Val(Mid(RutTemporal, 1, Len(RutTemporal) - 1))
  GrillaHonorarios.Col = 2
  PunteroTabla("Nombre") = Mid$(UCase(GrillaHonorarios.Text), 1, 30)
  GrillaHonorarios.Col = 3
  PunteroTabla("Otorgam") = Val(GrillaHonorarios.Text)
  GrillaHonorarios.Col = 4
  PunteroTabla("Novacion") = Val(GrillaHonorarios.Text)
End Sub

