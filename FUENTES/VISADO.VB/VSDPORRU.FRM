VERSION 4.00
Begin VB.Form VsdPorRut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultar Productos en Visado"
   ClientHeight    =   2040
   ClientLeft      =   8220
   ClientTop       =   4035
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   2445
   Left            =   8160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Top             =   3690
   Width           =   3600
   Begin Threed.SSFrame SSFrame1 
      Height          =   1980
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      _version        =   65536
      _extentx        =   5953
      _extenty        =   3493
      _stockprops     =   14
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
      Begin Threed.SSPanel SSPanel1 
         Height          =   690
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   3015
         _version        =   65536
         _extentx        =   5318
         _extenty        =   1217
         _stockprops     =   15
         backcolor       =   13160660
         Begin VB.TextBox Text1 
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   9.75
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   975
            MaxLength       =   13
            TabIndex        =   0
            Top             =   150
            Width           =   1665
         End

         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "RUT"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   12
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   225
            TabIndex        =   5
            Top             =   150
            Width           =   555
         End
      End

      Begin Threed.SSCommand Cancelar 
         Height          =   615
         Left            =   2325
         TabIndex        =   3
         Top             =   1200
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   3
         picture         =   "VSDPORRU.frx":0000
      End

      Begin Threed.SSCommand Aceptar 
         Height          =   615
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Aceptar"
         forecolor       =   8421376
         font3d          =   3
         picture         =   "VSDPORRU.frx":04D2
      End
   End
End
Attribute VB_Name = "VsdPorRut"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'
Private Sub Aceptar_Click()

Dim RutC As String
Dim DigV As String * 1

If Principal.PanelStatus.Visible Then
  Exit Sub
End If

Screen.MousePointer = 11
GRut_Continua = False

RutC = DesformateaRut(Text1.Text)
RutC = Mid$(RutC, 1, Len(RutC) - 1)
DigV = Right(Text1.Text, 1)

BaseVisado.Execute "Delete * From Vsd_Cli"

If ObtenerDigitoVerificador(RutC) = DigV Then ' RUT VALIDADO
  GRut_CForma = Text1.Text
  GRut_SForma = RutC + DigV
  Text1.Enabled = False
  'Aceptar.Enabled = False
  'Cancelar.Enabled = False

  If BajarDatosVisado_RUT Then
     GRut_Continua = True
     Unload Me
     Exit Sub
  ElseIf SePresionoTeclaESC Then
     Text1.Enabled = True
     'Aceptar.Enabled = True
     'Cancelar.Enabled = True
     Text1.SelStart = Len(Text1.Text)
     Text1.SetFocus
     GRut_Continua = False
     SePresionoTeclaESC = False
  Else
     MsgBox "No existen productos en visado para el cliente.", 16, "Visado de Productos"
     'Aceptar.Enabled = True
     'Cancelar.Enabled = True
     Text1.Enabled = True
     Text1.SelStart = Len(Text1.Text)
     Text1.SetFocus
     GRut_Continua = False
  End If
Else
  MsgBox "El Rut ingresado no es válido.", 16, "Visado de Productos"
  GRut_Continua = False
End If

Screen.MousePointer = 0

End Sub

Private Sub Aceptar_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    If Principal.Status.Visible Then
      SePresionoTeclaESC = True
    Else
      Unload Me
    End If

End If

End Sub


Private Sub Cancelar_Click()
  
  If Principal.Status.Visible Then
      SePresionoTeclaESC = True
    Else
      Unload Me
  End If
  GRut_Continua = False
  
End Sub


Private Sub Cancelar_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then

  If Principal.Status.Visible Then
      SePresionoTeclaESC = True
    Else
      Unload Me
  End If

End If

End Sub


Private Sub Form_Load()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub


Private Sub Text1_Click()
 
 If Trim(Text1.SelText) <> Trim(Text1.Text) Then
   Text1.SelStart = Len(Text1.Text)
   Text1.SetFocus
 End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Unload Me
    Exit Sub
End If

If Trim(Text1.SelText) = Trim(Text1.Text) Then
   Text1.Text = ""
End If
If KeyAscii = 13 And Text1.Text <> "" Then
   KeyAscii = 0
   Aceptar_Click
   Exit Sub
End If
If KeyAscii = 107 Then KeyAscii = 75
If ((KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 75) And (KeyAscii <> 8)) Then
  KeyAscii = 0
  Exit Sub
End If
If (InStr(Text1.Text, "K") <> 0) And (KeyAscii <> 8) Then
  KeyAscii = 0
  Exit Sub
End If

If KeyAscii = 8 Then
     Borrado$ = Right(Text1.Text, 1)
   If Len(Text1.Text) > 3 Then
     POS_SIGN$ = InStr(Text1.Text, "-")
     If POS_SIGN$ <> 0 Then
       ultimo$ = Mid(Text1.Text, POS_SIGN$ - 1, 1)
     End If
     Text1.Text = Format(Trim(Str(Mid(Text1.Text, 1, Len(Text1.Text) - 3))), "###,###,###") + "-" + ultimo$ + Borrado$
   ElseIf Len(Text1.Text) <= 3 And Len(Text1.Text) > 1 Then
     Text1.Text = Left(Text1.Text, 1) + Borrado$
   End If
   Text1.SelStart = Len(Text1.Text)
   Text1.SetFocus
   Exit Sub
End If

If Len(Text1.Text) > 0 And Len(Text1.Text) < 13 Then
 If InStr(Text1.Text, "-") <> 0 Then
   POS_SIGN$ = InStr(Text1.Text, "-")
   Text1.Text = Mid(Text1.Text, 1, POS_SIGN$ - 1) + Mid(Text1.Text, POS_SIGN$ + 1)
 End If
 If Len(Text1.Text) > 0 Then
   Text1.Text = Format(Text1.Text, "###,###,###-")
 End If
Else
 If Chr(KeyAscii) = "K" Or Len(Text1.Text) >= 13 Then
    KeyAscii = 0
 End If
End If

Text1.SelStart = Len(Text1.Text)
Text1.SetFocus

End Sub



Public Function BajarDatosVisado_RUT() As Boolean

Dim PunteroTabVsd As Recordset
Dim IncrementoStatus As Integer


If HayProductosVigentes(False) Then
    Screen.MousePointer = 11
    Set PunteroTabVsd = BaseVisado.OpenRecordset("Tab_Vsd", dbOpenTable)
    BaseVisado.Execute "DELETE * FROM Vsd_Prd;"
    BaseVisado.Execute "DELETE * FROM Vsd_Cli;"
    BajarDatosVisado_RUT = False
    PunteroTabVsd.MoveLast
    IncrementoStatus = Int(2615 / PunteroTabVsd.RecordCount)
    PunteroTabVsd.MoveFirst
    Principal.Status.Width = 15
    Principal.Status.Visible = True
    Principal.PanelStatus.Visible = True
    SePresionoTeclaESC = False
    While (Not PunteroTabVsd.EOF) And (Not SePresionoTeclaESC)
      If CorrespondeTraer(PunteroTabVsd("Vsd_Bajada")) Then
        If Mid$(PunteroTabVsd("Vsd_Vista_Ori"), 1, 2) <> "SY" Then
          PosicionRut = PunteroTabVsd("Vsd_Pos_Rut")
          LargoLlaveOrigen = PunteroTabVsd("Vsd_L_LLaveOri")
          PosicionProducto = PunteroTabVsd("Vsd_Pos_CodProd")
          Filtro = PunteroTabVsd("Vsd_Filtro")
          If Filtro Then
            Pos_Ftro = PunteroTabVsd("Vsd_Ftro_Pos")
            Dat_Ftro = PunteroTabVsd("Vsd_Ftro_Dat")
          End If
          Vista_Original = PunteroTabVsd("Vsd_Vista_Ori")
          If Mid$(PunteroTabVsd("Vsd_Vista_Ori"), 1, 2) <> "SY" Then
            If ObtieneInformacion(PunteroTabVsd("Vsd_Vista_Ori"), PunteroTabVsd("Vsd_Operacion")) Then
              BajarDatosVisado_RUT = True
            End If
          End If
        End If
      End If
      PunteroTabVsd.MoveNext
      Principal.Status.Width = Principal.Status.Width + IncrementoStatus
    Wend
    PunteroTabVsd.Close
    Principal.Status.Width = 2615
    Principal.Status.Visible = False
    Principal.PanelStatus.Visible = False
    If SePresionoTeclaESC Then
      VerMensaje = False
      BajarDatosVisado_RUT = False
    End If
    Call LimpiezaVisado
    Screen.MousePointer = 0
Else
    MsgBox "Debe Seleccionar Productos a Visar", 0 + 16, "Falta Vigencia Productos"
End If
  
End Function


Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

Dim RutVsd As String
Dim RutPvta As String
Dim TexOb As String

If KeyCode = vbKeyF3 Then
  RutVsd = SacaValorINI("Parametros", "Param1", DIRECTORIO_01 + ARCHIVO_INI01)
  RutPvta = SacaValorINI("ULTLLAVES", "ULTRUT", "C:\DATA\POSVENTA\GETKEY.INI")
  TexOb = Text1.Text
  
  RutVsd = IIf(Trim$(RutVsd) <> "", Right(String(10, "0") + Trim$(RutVsd), 10), "")
  RutPvta = IIf(Trim$(RutPvta) <> "", Right(String(10, "0") + Trim$(RutPvta), 10), "")
  TexOb = IIf(Trim$(TexOb) <> "", Right(String(10, "0") + DesformateaRut(TexOb), 10), "")
  
  If (RutVsd <> "" And TexOb = RutPvta) Or (RutVsd <> "" And Text1.Text = "") Then
    Text1.Text = FormateaRut(RutVsd)
  ElseIf (RutPvta <> "" And TexOb = RutVsd) Or (RutPvta <> "" And Text1.Text = "") Then
    Text1.Text = FormateaRut(RutPvta)
  End If
  
  Text1.SelStart = Len(Text1.Text)
  
End If

End Sub


