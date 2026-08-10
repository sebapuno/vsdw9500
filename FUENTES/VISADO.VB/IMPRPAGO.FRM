VERSION 4.00
Begin VB.Form ImprimePagoNotario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión Pagos Notario"
   ClientHeight    =   2535
   ClientLeft      =   1665
   ClientTop       =   2190
   ClientWidth     =   3705
   ControlBox      =   0   'False
   Height          =   2940
   Left            =   1605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Top             =   1845
   Width           =   3825
   Begin Threed.SSFrame SSFrame1 
      Height          =   2490
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   3630
      _version        =   65536
      _extentx        =   6403
      _extenty        =   4392
      _stockprops     =   14
      caption         =   "Ingreso Rango de Fecha ..."
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
      Begin Threed.SSPanel SSPanel7 
         Height          =   570
         Left            =   225
         TabIndex        =   5
         Top             =   360
         Width           =   3165
         _version        =   65536
         _extentx        =   5583
         _extenty        =   1005
         _stockprops     =   15
         forecolor       =   16711680
         backcolor       =   -2147483644
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   7.5
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox FechaDesde 
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   400
               size            =   9.75
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            TabIndex        =   9
            Text            =   "19990520"
            Top             =   120
            Width           =   1275
         End

         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha &desde"
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
            Height          =   240
            Left            =   380
            TabIndex        =   0
            Top             =   165
            Width           =   1200
         End
      End

      Begin Threed.SSPanel SSPanel1 
         Height          =   570
         Left            =   225
         TabIndex        =   6
         Top             =   1005
         Width           =   3165
         _version        =   65536
         _extentx        =   5583
         _extenty        =   1005
         _stockprops     =   15
         forecolor       =   16711680
         backcolor       =   -2147483644
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   7.5
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox FechaHasta 
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   400
               size            =   9.75
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1650
            TabIndex        =   8
            Text            =   "19990520"
            Top             =   120
            Width           =   1275
         End

         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha &hasta"
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
            Height          =   240
            Left            =   380
            TabIndex        =   1
            Top             =   165
            Width           =   1110
         End
      End

      Begin VB.Label Label3 
         Caption         =   "Formato : aaaammdd"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   1920
         Width           =   1485
      End

      Begin Threed.SSCommand Imprimir 
         Height          =   645
         Left            =   1800
         TabIndex        =   2
         Top             =   1710
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Imprimir"
         forecolor       =   8388608
         font3d          =   1
         autosize        =   1
         picture         =   "IMPRPAGO.frx":0000
      End

      Begin Threed.SSCommand Cancelar 
         Height          =   645
         Left            =   2610
         TabIndex        =   3
         Top             =   1710
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   1
         picture         =   "IMPRPAGO.frx":031A
      End
   End
End
Attribute VB_Name = "ImprimePagoNotario"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Cancelar_Click()
  Unload Me
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub FechaDesde_GotFocus()
  Label1.ForeColor = COLOR_09
  Label1.FontUnderline = True
End Sub
Private Sub FechaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  If (Shift = 0) And ((KeyCode = 38) Or (KeyCode = 40)) Then
    FechaHasta.SetFocus
  End If
End Sub
Private Sub FechaDesde_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      FechaHasta.SetFocus
    Case 27
      Unload Me
  End Select
End Sub

Private Sub FechaDesde_LostFocus()
  Label1.ForeColor = COLOR_10
  Label1.FontUnderline = False
End Sub

Private Sub FechaHasta_GotFocus()
  Label2.ForeColor = COLOR_09
  Label2.FontUnderline = True
End Sub
Private Sub FechaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  If (Shift = 0) And ((KeyCode = 38) Or (KeyCode = 40)) Then
    FechaDesde.SetFocus
  End If
End Sub
Private Sub FechaHasta_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Imprimir.SetFocus
    Case 27
      Unload Me
  End Select
End Sub

Private Sub FechaHasta_LostFocus()
  Label2.ForeColor = COLOR_10
  Label2.FontUnderline = False
End Sub

Private Sub Form_Activate()
  FechaDesde.Text = Format(Now, "yyyymmdd")
  FechaHasta.Text = Format(Now, "yyyymmdd")
  Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
  Screen.MousePointer = 11
  Principal.Enabled = False
  Call CentrarVentana(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Principal.Enabled = True
End Sub
Private Sub Imprimir_Click()
Dim Retorno As Integer
Dim FechaFinal As String
Dim FechaInicial As String
Dim NombreArchivo As String
  If FechaDesde.Text <> "" Then
    If FechaHasta.Text <> "" Then
      If Val(FechaDesde.Text) <= Val(FechaHasta.Text) Then
        NombreArchivo = DIRECTORIO_02 + ARCHIVO_EXE05
        If Dir(NombreArchivo) = UCase(Mid(NombreArchivo, Len(NombreArchivo) - 10)) Then
          Screen.MousePointer = 11
          'FechaFinal = Format(FechaHasta.Text, "ddmmyy")
          'FechaInicial = Format(FechaDesde.Text, "ddmmyy")
          FechaFinal = FechaHasta.Text
          FechaInicial = FechaDesde.Text
          Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
          Retorno = PoneValorINI("Parametros", "Param1", FechaInicial, DIRECTORIO_01 + ARCHIVO_INI01)
          Retorno = PoneValorINI("Parametros", "Param2", FechaFinal, DIRECTORIO_01 + ARCHIVO_INI01)
          Call EjecutarPrograma(NombreArchivo, 0)
          Screen.MousePointer = 1
        Else
          MsgBox "No se Encontró Programa " + NombreArchivo, 0 + 16, "Falta Programa para Imprimir"
        End If
      Else
        MsgBox "Fechas Intercambiadas", 0 + 16, "Datos Incorrectos"
        FechaHasta.SetFocus
      End If
    Else
      MsgBox "Falta Ingresar Fecha Final", 0 + 16, "Faltan Datos"
      FechaHasta.SetFocus
    End If
  Else
    MsgBox "Falta Ingresar Fecha Inicial", 0 + 16, "Faltan Datos"
    FechaDesde.SetFocus
  End If
End Sub
Private Sub Imprimir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub
Private Sub Label1_Click()
  FechaDesde.SetFocus
End Sub
Private Sub Label2_Click()
  FechaHasta.SetFocus
End Sub
