VERSION 4.00
Begin VB.Form ConfiguraOficina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modificación Oficina a Visar"
   ClientHeight    =   3165
   ClientLeft      =   2505
   ClientTop       =   1770
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Height          =   3570
   Left            =   2445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Top             =   1425
   Width           =   5295
   Begin Threed.SSFrame SSFrame1 
      Height          =   3105
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5100
      _version        =   65536
      _extentx        =   8996
      _extenty        =   5477
      _stockprops     =   14
      caption         =   "Selección de Oficina ..."
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
      Begin VB.ComboBox ListaOficinas 
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
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   4845
      End
      Begin Threed.SSFrame Frame3D1 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   4845
         _version        =   65536
         _extentx        =   8546
         _extenty        =   1931
         _stockprops     =   14
         caption         =   "Mantener Productos de hasta :"
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
         Begin VB.OptionButton Option1 
            Caption         =   "15 Días"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   9
            Top             =   330
            Width           =   1125
         End
         Begin VB.OptionButton Option2 
            Caption         =   "30 Días"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   630
            Width           =   1125
         End
         Begin VB.OptionButton Option3 
            Caption         =   "60 Días"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2910
            TabIndex        =   7
            Top             =   330
            Width           =   1125
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Inhabilitado"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2910
            TabIndex        =   6
            Top             =   630
            Width           =   1425
         End
      End
      Begin Threed.SSPanel Modo 
         Height          =   330
         Left            =   180
         TabIndex        =   4
         Top             =   2610
         Width           =   1410
         _version        =   65536
         _extentx        =   2487
         _extenty        =   582
         _stockprops     =   15
         forecolor       =   -2147483640
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Courier New"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         font3d          =   3
      End
      Begin Threed.SSCommand Cancelar 
         Height          =   615
         Left            =   4170
         TabIndex        =   3
         Top             =   2355
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   3
         picture         =   "OFICINA.frx":0000
      End
      Begin Threed.SSCommand Aceptar 
         Height          =   615
         Left            =   3270
         TabIndex        =   2
         Top             =   2355
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Aceptar"
         forecolor       =   8421376
         font3d          =   3
         picture         =   "OFICINA.frx":04D2
      End
   End
End
Attribute VB_Name = "ConfiguraOficina"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Respuesta As Integer

Private Sub Aceptar_Click()

If ListaOficinas.ListCount > 0 Then
  I% = MsgBox("¿Seguro de Cambiar la Oficina?", 4 + 32 + 256, "Confirmar Modificación")
  If I% = 6 Then
    Oficina = Mid$(ListaOficinas.Text, 1, 3)
    Respuesta = PoneValorINI("Visado", "Oficina", Oficina, DIRECTORIO_01 + ARCHIVO_INI01)
    Select Case True
           Case Option1.Value
              Respuesta = PoneValorINI("Visado", "Dias", "15", DIRECTORIO_01 + ARCHIVO_INI01)
           Case Option2.Value
              Respuesta = PoneValorINI("Visado", "Dias", "30", DIRECTORIO_01 + ARCHIVO_INI01)
           Case Option3.Value
              Respuesta = PoneValorINI("Visado", "Dias", "60", DIRECTORIO_01 + ARCHIVO_INI01)
           Case Option4.Value
              Respuesta = PoneValorINI("Visado", "Dias", "99", DIRECTORIO_01 + ARCHIVO_INI01)
    End Select
    Unload Me
  End If
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
Private Sub Form_Activate()

If ListaOficinas.ListCount > 0 Then
  ListaOficinas.ListIndex = UbicacionPosicionOficina()
End If
Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
Dim Dias As String * 2
Dim BaseOficina As Database
Dim PunteroOficina As Recordset
Dim OficinaEncontrada As Boolean

  Screen.MousePointer = 11
  
  OficinaEncontrada = False
    
  Principal.Enabled = False
  Call CentrarVentana(Me)
  
  If InStr("00,54,55,56", Environ$(NODOMETRO)) = 0 Then
    Modo.Caption = "Regional"
  Else
    Modo.Caption = "Metro"
  End If
  
  Set BaseOficina = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
  Set PunteroOficina = BaseOficina.OpenRecordset("SELECT * FROM TBLCUI WHERE CUI_CUISUC=" + Environ("Oficina") + " AND CUI_CUIISU='S'", dbOpenDynaset)
  While Not PunteroOficina.EOF
    ListaOficinas.AddItem GlosaOficina(Format(PunteroOficina!CUI_CUICOD, "000"))
    
    If PunteroOficina!CUI_CUICOD = Oficina Then
      OficinaEncontrada = True
    End If
    PunteroOficina.MoveNext
  Wend
  PunteroOficina.Close
  BaseOficina.Close

  Set BaseOficina = OpenDatabase(DIRECTORIO_01 + "OFICINA.MDB")
  Set PunteroOficina = BaseOficina.OpenRecordset("SELECT * FROM grp_ofi WHERE grp_oficui='" + Environ("Oficina") + "'", dbOpenDynaset)
  While Not PunteroOficina.EOF
    ListaOficinas.AddItem GlosaOficina(Format(PunteroOficina!GRP_OFICOD, "000"))
    
    If PunteroOficina!GRP_OFICOD = Oficina Then
      OficinaEncontrada = True
    End If
    
    PunteroOficina.MoveNext
  Wend
  PunteroOficina.Close
  BaseOficina.Close
  
  Dias = SacaValorINI("Visado", "Dias", DIRECTORIO_01 + ARCHIVO_INI01)
  
  Select Case Dias
   Case "15"
      Option1.Value = True
   Case "30"
      Option2.Value = True
   Case "60"
      Option3.Value = True
   Case Else
      Option4.Value = True
  End Select

  If Not OficinaEncontrada Then
    Oficina = Environ("OFICINA")
  End If

  For NOfi% = 0 To ListaOficinas.ListCount - 1
    If Mid$(ListaOficinas.List(NOfi%), 1, 3) = Oficina Then
       ListaOficinas.ListIndex = NOfi%
    End If
  Next NOfi%
  
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Principal.Enabled = True
End Sub

Private Sub ListaOficinas_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Aceptar_Click
    Case 27
      Unload Me
  End Select
End Sub


Private Function UbicacionPosicionOficina() As Integer
Dim I As Integer

UbicacionPosicionOficina = 0
For I = 0 To ListaOficinas.ListCount - 1
    ListaOficinas.ListIndex = I
    If Mid$(ListaOficinas.Text, 1, 1) = "(" Then
      If Oficina = Mid$(ListaOficinas.Text, 1, 3) Then
        UbicacionPosicionOficina = I
        Exit For
      End If
    Else
      If Oficina = Mid$(ListaOficinas.Text, 1, 3) Then
        UbicacionPosicionOficina = I
        Exit For
      End If
    End If
Next I

End Function

