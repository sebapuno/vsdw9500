VERSION 4.00
Begin VB.Form Principal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Excepciones de Visado"
   ClientHeight    =   4245
   ClientLeft      =   1155
   ClientTop       =   1485
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   4650
   Left            =   1095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Top             =   1140
   Width           =   7620
   Begin Threed.SSPanel Panel3D1 
      Height          =   3000
      Left            =   4770
      TabIndex        =   31
      Top             =   495
      Width           =   2670
      _version        =   65536
      _extentx        =   4710
      _extenty        =   5292
      _stockprops     =   15
      forecolor       =   16384
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelouter      =   0
      bevelinner      =   2
      shadowcolor     =   1
      font3d          =   1
      alignment       =   6
      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   7
         ItemData        =   "PTACRN04.frx":0000
         Left            =   1170
         List            =   "PTACRN04.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2610
         Width           =   1365
      End

      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   6
         ItemData        =   "PTACRN04.frx":0004
         Left            =   1170
         List            =   "PTACRN04.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2250
         Width           =   1365
      End

      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   5
         ItemData        =   "PTACRN04.frx":0008
         Left            =   1170
         List            =   "PTACRN04.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1890
         Width           =   1365
      End

      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   4
         ItemData        =   "PTACRN04.frx":000C
         Left            =   1170
         List            =   "PTACRN04.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1530
         Width           =   1365
      End

      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   3
         ItemData        =   "PTACRN04.frx":0010
         Left            =   1170
         List            =   "PTACRN04.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1170
         Width           =   1365
      End

      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   2
         ItemData        =   "PTACRN04.frx":0014
         Left            =   1170
         List            =   "PTACRN04.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   810
         Width           =   1365
      End

      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   1
         ItemData        =   "PTACRN04.frx":0018
         Left            =   1170
         List            =   "PTACRN04.frx":001A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   450
         Width           =   1365
      End

      Begin VB.ComboBox VistoBueno 
         Height          =   300
         Index           =   0
         ItemData        =   "PTACRN04.frx":001C
         Left            =   1170
         List            =   "PTACRN04.frx":001E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   90
         Width           =   1365
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   7
         Height          =   250
         Left            =   90
         TabIndex        =   22
         Top             =   2620
         Width           =   980
         Text            =   "28/01/1997"
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   6
         Height          =   250
         Left            =   105
         TabIndex        =   19
         Top             =   2260
         Width           =   980
         Text            =   "28/01/1997"
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   5
         Height          =   250
         Left            =   105
         TabIndex        =   16
         Top             =   1900
         Width           =   980
         Text            =   "28/01/1997"
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   4
         Height          =   250
         Left            =   90
         TabIndex        =   13
         Top             =   1540
         Width           =   980
         Text            =   "28/01/1997"
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   3
         Height          =   250
         Left            =   90
         TabIndex        =   10
         Top             =   1180
         Width           =   980
         Text            =   "28/01/1997"
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   2
         Height          =   250
         Left            =   90
         TabIndex        =   7
         Top             =   820
         Width           =   980
         Text            =   "28/01/1997"
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   1
         Height          =   250
         Left            =   90
         TabIndex        =   4
         Top             =   460
         Width           =   980
         Text            =   "28/01/1997"
      End
      ' FI2026 - VBX.DateTextBox migrado a VB.TextBox nativo
      Begin VB.TextBox FechaCompromiso 
         Index           =   0
         Height          =   250
         Left            =   90
         TabIndex        =   1
         Top             =   100
         Width           =   980
         Text            =   "28/01/1997"
      End
   End

   Begin VB.CheckBox Normas 
      Caption         =   "Certificación &Notarial de Pagaré en Orden."
      Height          =   250
      Index           =   7
      Left            =   90
      TabIndex        =   21
      Top             =   3150
      Width           =   4600
   End

   Begin VB.CheckBox Normas 
      Caption         =   "&Garantías Cumplen con las Normas Exigidas."
      Height          =   250
      Index           =   6
      Left            =   90
      TabIndex        =   18
      Top             =   2790
      Width           =   4600
   End

   Begin VB.CheckBox Normas 
      Caption         =   "Cliente Creado y con Ejecutivo Control en C.&D.R."
      Height          =   250
      Index           =   5
      Left            =   90
      TabIndex        =   15
      Top             =   2430
      Width           =   4600
   End

   Begin VB.CheckBox Normas 
      Caption         =   "&Línea de Crédito Vigente y Margen Disponible."
      Height          =   250
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   2070
      Width           =   4600
   End

   Begin VB.CheckBox Normas 
      Caption         =   "&M.A.C. o Solicitud con Atribuciones y Facultades Suficientes."
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   1710
      Width           =   4600
   End

   Begin VB.CheckBox Normas 
      Caption         =   "Pagaré Bien &Extendido y Consistente."
      Height          =   250
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   1350
      Width           =   4600
   End

   Begin VB.CheckBox Normas 
      Caption         =   "&Vigencia de Sociedad."
      Height          =   250
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   4600
   End

   Begin VB.CheckBox Normas 
      Caption         =   "&Prohibición y Márgenes Articulos 64 y 65."
      ForeColor       =   &H00000000&
      Height          =   250
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   4600
   End

   Begin Threed.SSCommand Cancelar 
      Height          =   615
      Left            =   6570
      TabIndex        =   25
      Top             =   3555
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      font3d          =   3
      picture         =   "PTACRN04.frx":0020
   End

   Begin Threed.SSCommand Aceptar 
      Height          =   615
      Left            =   5625
      TabIndex        =   24
      Top             =   3555
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Aceptar"
      forecolor       =   8421376
      font3d          =   3
      picture         =   "PTACRN04.frx":0832
   End

   Begin VB.Label Encabezado 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Visto Bueno"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   5925
      TabIndex        =   30
      Top             =   135
      Width           =   1515
   End

   Begin VB.Label Encabezado 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Normas a Cumplir"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   45
      TabIndex        =   29
      Top             =   135
      Width           =   4710
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
      Height          =   420
      Index           =   2
      Left            =   5925
      TabIndex        =   28
      Top             =   45
      Width           =   1515
   End

   Begin VB.Label Encabezado 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Fecha Comprometida"
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   4780
      TabIndex        =   27
      Top             =   45
      Width           =   1110
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
      Height          =   420
      Index           =   0
      Left            =   45
      TabIndex        =   26
      Top             =   45
      Width           =   4710
   End
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Const HWND_TOPMOST = -1
Const SWP_NOSIZENOMOVE = 3

Const ARCHIVO_INI01 = "Visado.INI"
Const DIRECTORIO_01 = "C:\Data\Visado\"
Const DIRECTORIO_02 = "C:\Data\PilExc\"
Const ARCHIVO_DAT01 = "Autoriza.DAT"
Const ARCHIVO_MDB01 = "Visado.MDB"
Const MARCAEXCEPCIONVISADOR = "3"

Dim Datos As String
Dim RutCliente As String
Dim Correlativo As Integer
Dim BaseVisado As Database
Dim CorrelativoDisponible As Integer
Private Sub Aceptar_Click()
Dim I As Integer
Dim Rut As String
Dim PunteroVsdPrd As Recordset
  If SeIngresaronExcepciones Then
    If TodosLosCamposLlenos Then
      I = MsgBox("¿ Son Correctas las Excepciones ?", 4 + 32 + 256, "Verificar Excepciones")
      If I = 6 Then
        For I = 0 To 7
          If (Normas(I).Enabled) And (Normas(I).Value = 1) Then
            Rut = DesformateaRut(VistoBueno(I).Text)
            Select Case I
              Case 0
                Datos = Mid$(Datos, 1, 26) + MARCAEXCEPCIONVISADOR + Mid$(Datos, 27, 1) + Rut + _
                              String(30 - Len(Rut), " ") + Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Mid$(Datos, 67)
              Case 1
                Datos = Mid$(Datos, 1, 27) + MARCAEXCEPCIONVISADOR + Rut + String(30 - Len(Rut), " ") + _
                              Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Mid$(Datos, 67)
              Case 2
                Datos = Mid$(Datos, 1, 121) + MARCAEXCEPCIONVISADOR + Rut + String(30 - Len(Rut), " ") + _
                              Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Mid$(Datos, 161)
              Case 3
                Datos = Mid$(Datos, 1, 215) + MARCAEXCEPCIONVISADOR + Rut + String(30 - Len(Rut), " ") + _
                              Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Mid$(Datos, 255)
              Case 4
                Datos = Mid$(Datos, 1, 262) + String(2, MARCAEXCEPCIONVISADOR) + Rut + _
                              String(30 - Len(Rut), " ") + Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Mid$(Datos, 303)
              Case 5
                Datos = Mid$(Datos, 1, 357) + MARCAEXCEPCIONVISADOR + Rut + String(30 - Len(Rut), " ") + _
                              Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Mid$(Datos, 397)
              Case 6
                Datos = Mid$(Datos, 1, 546) + String(6, MARCAEXCEPCIONVISADOR) + Rut + _
                              String(30 - Len(Rut), " ") + Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Mid$(Datos, 591)
              Case 7
                Datos = Mid$(Datos, 1, 794) + MARCAEXCEPCIONVISADOR + _
                              Format$(FechaCompromiso(I).Text, "ddmmyyyy") + Rut
            End Select
          End If
        Next I
        Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
        ' FI2026 - era: PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name
        ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
        ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
        ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
        PunteroVsdPrd.Index = BaseVisado.TableDefs("Vsd_Prd").Indexes(0).Name
        PunteroVsdPrd.Seek "=", RutCliente, "EXC", Correlativo
        If PunteroVsdPrd.NoMatch Then
          PunteroVsdPrd.AddNew
          PunteroVsdPrd("Prd_Data") = Datos
          PunteroVsdPrd("Prd_Correl") = CorrelativoDisponible
          PunteroVsdPrd("Prd_RutCli") = RutCliente
          PunteroVsdPrd("Prd_CodProd") = "EXC"
        Else
          PunteroVsdPrd.Edit
          PunteroVsdPrd("Prd_Data") = Datos
        End If
        PunteroVsdPrd.Update
        PunteroVsdPrd.Close
        Retorno% = PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01)
        End
      End If
    Else
      MsgBox "Se Detectó Excepción Incompleta", 0 + 16, "Faltan Datos"
    End If
  Else
    MsgBox "No se Han Ingresado Excepciones", 0 + 16, "Faltan Datos"
  End If
End Sub
Private Sub Aceptar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub Cancelar_Click()
  End
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub FechaCompromiso_GotFocus(Index As Integer)
  Call ColorearNorma(&H800080, Index, True)
  FechaCompromiso(Index).ForeColor = &HFF&
End Sub

Private Sub FechaCompromiso_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 27 Then
      Call Cancelar_Click
  End If
End Sub
Private Sub FechaCompromiso_LostFocus(Index As Integer)
Dim mmdd As String * 4
  
  mmdd = Mid$(FechaCompromiso(Index).Text, 4, 2) + Mid$(FechaCompromiso(Index).Text, 1, 2)
  If WeekDay(CDate(FechaCompromiso(Index).Text)) = 7 Or WeekDay(CDate(FechaCompromiso(Index).Text)) = 1 Then
     MsgBox " La fecha Ingresada no es Dia Hábil  "
     FechaCompromiso(Index).SetFocus
  Else
     If InStr("0101,0328,0329,0501,0521,0529,0815,0911,0918,0919,1101,1208,1225,1231", mmdd) > 0 Then
        MsgBox " La fecha Ingresada es Feriado  "
        FechaCompromiso(Index).SetFocus
     Else
        Call ColorearNorma(&H0&, Index, False)
        FechaCompromiso(Index).ForeColor = &H0&
     End If
  End If
End Sub

Private Sub Form_Activate()
  Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
Dim I As Integer
Dim Oficina As String * 3
Dim Largo As Integer
Dim TieneEXC As Integer
Dim Posicion As Integer
Dim Producto As String
Dim LineaDatos As String
Dim NombreCliente  As String * 30
Dim FiltroVsdPrd   As Recordset
Dim FiltroTabVisar As Recordset
Dim PatronProductoPrincipal As String
Dim PatronProductoRelacionado As String
  TieneEXC = False
  Screen.MousePointer = 11
  Call CentrarVentana(Me)
  Sucess% = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZENOMOVE)
  Correlativo = -1
  CorrelativoDisponible = 1
  For I = 0 To 7
    VistoBueno(I).ListIndex = -1
    FechaCompromiso(I).Text = Format(Now, "dd/mm/yyyy")
  Next I
  Producto = SacaValorINI("Parametros", "Param2", DIRECTORIO_01 + ARCHIVO_INI01)
  RutCliente = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  PatronProductoPrincipal = SacaValorINI("Parametros", "LlaveVisad", DIRECTORIO_01 + ARCHIVO_INI01)
  Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
  Set FiltroTabVisar = BaseVisado.OpenRecordset("SELECT * FROM Tab_Visar WHERE Vis_CodProd='" + Producto + "' AND Vis_CodRel='EXC';")
  Oficina = Environ$("OFICINA")
  Largo = Interpreta(FiltroTabVisar("Vis_PosProd"), "L")
  Posicion = Interpreta(FiltroTabVisar("Vis_PosProd"), "P")
  FiltroTabVisar.Close
  Datos = Mid$(PatronProductoPrincipal, 1, 3) + Mid$(PatronProductoPrincipal, Posicion + 3, Largo - 3)
  Datos = Datos + String(619, "0") + Format$(Date, "ddmmyyyy") + Mid$(PatronProductoPrincipal, 27, 3) + _
                Mid$(PatronProductoPrincipal, 4, 3) + Mid$(PatronProductoPrincipal, 4, 3) + _
                Mid$(PatronProductoPrincipal, 7, 3) + Mid$(PatronProductoPrincipal, 7, 3) + _
                String(30, " ") + Oficina + Mid$(PatronProductoPrincipal, 61, 17) + _
                Mid$(PatronProductoPrincipal, 26, 1) + Mid$(PatronProductoPrincipal, 39, 3) + String(91, " ")
  PatronProductoPrincipal = Mid$(PatronProductoPrincipal, Posicion + 3, Largo - 3)
  Set FiltroVsdPrd = BaseVisado.OpenRecordset("SELECT * FROM Vsd_Prd WHERE Prd_RutCli='" + _
                                                                                 RutCliente + "' AND Prd_CodProd='EXC' ORDER BY Prd_Correl;", dbOpenSnapshot)
  If FiltroVsdPrd.RecordCount > 0 Then
    Do While Not FiltroVsdPrd.EOF
      Set FiltroTabVisar = BaseVisado.OpenRecordset("SELECT * FROM Tab_Visar WHERE Vis_CodProd='" + Producto + "' AND Vis_CodRel='EXC';", dbOpenSnapshot)
      Largo = Interpreta(FiltroTabVisar("Vis_PosRel"), "L")
      Posicion = Interpreta(FiltroTabVisar("Vis_PosRel"), "P")
      FiltroTabVisar.Close
      PatronProductoRelacionado = Mid$(FiltroVsdPrd("Prd_Data"), Posicion + 3, Largo - 3)
      If PatronProductoPrincipal = PatronProductoRelacionado Then
        TieneEXC = True
        Datos = FiltroVsdPrd("Prd_Data")
        Correlativo = FiltroVsdPrd("Prd_Correl")
        If Mid$(Datos, 27, 1) <> "0" Then
          Call AsignarValorNorma(Datos, 0, 29, 59)
        End If
        If Mid$(Datos, 28, 1) <> "0" Then
          Call AsignarValorNorma(Datos, 1, 29, 59)
        End If
        If Mid$(Datos, 122, 1) <> "0" Then
          Call AsignarValorNorma(Datos, 2, 123, 153)
        End If
        If Mid$(Datos, 216, 1) <> "0" Then
          Call AsignarValorNorma(Datos, 3, 217, 247)
        End If
        If (Mid$(Datos, 263, 1) <> "0") Or (Mid$(Datos, 264, 1) <> "0") Then
          Call AsignarValorNorma(Datos, 4, 265, 295)
        End If
        If Mid$(Datos, 358, 1) <> "0" Then
          Call AsignarValorNorma(Datos, 5, 359, 389)
        End If
        If (Mid$(Datos, 547, 1) <> "0") Or (Mid$(Datos, 548, 1) <> "0") Or _
           (Mid$(Datos, 549, 1) <> "0") Or (Mid$(Datos, 550, 1) <> "0") Or _
           (Mid$(Datos, 551, 1) <> "0") Or (Mid$(Datos, 552, 1) <> "0") Then
          Call AsignarValorNorma(Datos, 6, 553, 583)
        End If
        If Mid$(Datos, 795, 1) = "3" Then
           Call AsignarValorNorma(Datos, 7, 804, 796)
        End If
        Exit Do
      End If
      FiltroVsdPrd.MoveNext
    Loop
    FiltroVsdPrd.MoveLast
    CorrelativoDisponible = FiltroVsdPrd("Prd_Correl") + 1
  End If
  FiltroVsdPrd.Close
  If Not TieneEXC Then
     NombreCliente = Trae_Nombre(Mid(Datos, 10, 9), z)
     Datos = Mid$(Datos, 1, 668) + NombreCliente + Mid$(Datos, 699, 96)
  End If
  If Dir(DIRECTORIO_02 + ARCHIVO_DAT01) = UCase(ARCHIVO_DAT01) Then
    Open DIRECTORIO_02 + ARCHIVO_DAT01 For Input As #1
    While Not EOF(1)
      Line Input #1, LineaDatos
      For I = 0 To 7
        If LTrim(Mid$(LineaDatos, 1, 9)) <> "" And Val(Mid$(LineaDatos, 1, 8)) <> 0 Then
           If VistoBueno(I).Enabled Then
             VistoBueno(I).AddItem FormateaRut(Mid$(LineaDatos, 1, 9))
             VistoBueno(I).ListIndex = 0
           End If
        Else
           MsgBox "Debe antes ingresar los Ruts Autorizadores" + Chr(13) + _
                  "por Mantención de Excepciones"
           End
        End If
      Next I
    Wend
    Close 1
  Else
    MsgBox "No se Encontró Archivo" + Chr(10) + DIRECTORIO_02 + ARCHIVO_DAT01 + Chr(13) + _
           "Debe Cargar el Producto " + Chr(13) + " Mantención de Excepciones", 0 + 16, "PTACRE04"
    End
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  BaseVisado.Close
End Sub
Private Sub Normas_GotFocus(Index As Integer)
  Call ColorearNorma(&H800080, Index, True)
End Sub

Private Sub Normas_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Aceptar_Click
    Case 27
      Call Cancelar_Click
  End Select
End Sub
Private Sub Normas_LostFocus(Index As Integer)
  Call ColorearNorma(&H0&, Index, False)
End Sub
Private Sub ColorearNorma(ByVal Color As Long, ByVal Indice As Integer, ByVal Estado As Boolean)
  Normas(Indice).ForeColor = Color
End Sub
Private Sub CentrarVentana(Forma As Form)
  Forma.Left = (Screen.Width - Forma.Width) / 2
  Forma.Top = (Screen.Height - Forma.Height) / 2
End Sub
Private Function Interpreta(Cadena As String, DatoBuscado As String) As Integer
Dim PosicionComa As Integer
  PosicionComa = InStr(Cadena, ",")
  If PosicionComa <> 0 Then
    If DatoBuscado = "P" Then
      Interpreta = Val(Mid(Cadena, 1, PosicionComa - 1))
    Else
      Interpreta = Val(Mid(Cadena, PosicionComa + 1, Len(Cadena)))
    End If
  Else
    Interpreta = -1
  End If
End Function
Private Sub VistoBueno_GotFocus(Index As Integer)
  Call ColorearNorma(&H800080, Index, True)
End Sub

Private Sub VistoBueno_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Aceptar_Click
    Case 27
      Call Cancelar_Click
  End Select
End Sub
Private Sub VistoBueno_LostFocus(Index As Integer)
  Call ColorearNorma(&H0&, Index, False)
End Sub

Private Sub AsignarValorNorma(ByVal Data As String, ByVal Indice As Integer, ByVal PosicionRut As Integer, ByVal PosicionFecha As Integer)
Dim Rut As String
  If InStr("0123456789K", Mid$(Data, PosicionRut + 9, 1)) <> 0 Then
    Rut = LTrim$(Str(Val(Mid(Data, PosicionRut, 9)))) + Mid$(Data, PosicionRut + 9, 1)
  Else
    Rut = LTrim$(Str(Val(Mid(Data, PosicionRut, 10))))
  End If
  Normas(Indice).Value = 1
  VistoBueno(Indice).Clear
  VistoBueno(Indice).AddItem FormateaRut(Rut)
  VistoBueno(Indice).ListIndex = 0
  FechaCompromiso(Indice).Text = Format(Mid(Data, PosicionFecha, 8), "&&/&&/&&&&")
  Normas(Indice).Enabled = False
  VistoBueno(Indice).Enabled = False
  FechaCompromiso(Indice).Enabled = False
End Sub
Private Function DesformateaRut(ByVal Rut As String) As String
Dim Posicion As Integer
Dim NumeroRut As String
  Posicion = 1
  NumeroRut = ""
  While Posicion <= Len(Rut)
    If (Mid$(Rut, Posicion, 1) <> ".") And (Mid$(Rut, Posicion, 1) <> "-") Then
      NumeroRut = NumeroRut + Mid$(Rut, Posicion, 1)
    End If
    Posicion = Posicion + 1
  Wend
  DesformateaRut = String(10 - Len(NumeroRut), "0") + NumeroRut
End Function
Private Function FormateaRut(Rut As String) As String
  Rut = String(9 - Len(Rut), "0") + Rut
  FormateaRut = Mid$(Rut, 1, 2) + "." + Mid$(Rut, 3, 3) + "." + Mid$(Rut, 6, 3) + "-" + UCase$(Mid$(Rut, 9, 1))
End Function

Private Function TodosLosCamposLlenos() As Boolean
Dim I As Integer
  TodosLosCamposLlenos = True
  For I = 0 To 7
    If (Normas(I).Enabled) And (Normas(I).Value = 1) Then
      If (Not IsDate(FechaCompromiso(I).Text)) Or (Trim$(VistoBueno(I).Text) = "") Then
        TodosLosCamposLlenos = False
        Exit For
      End If
    End If
  Next I
End Function


Private Function SeIngresaronExcepciones() As Boolean
Dim I As Integer
  SeIngresaronExcepciones = False
  For I = 0 To 7
    If (Normas(I).Enabled) And (Normas(I).Value = 1) Then
      SeIngresaronExcepciones = True
      Exit For
    End If
  Next I
End Function
