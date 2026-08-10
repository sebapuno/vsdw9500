VERSION 4.00
Begin VB.Form RepararArchivos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparar Bases de Datos"
   ClientHeight    =   2700
   ClientLeft      =   3615
   ClientTop       =   3165
   ClientWidth     =   2865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   3105
   Left            =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   2865
   ShowInTaskbar   =   0   'False
   Top             =   2820
   Width           =   2985
   Begin Threed.SSFrame SSFrame1 
      Height          =   1890
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   2745
      _version        =   65536
      _extentx        =   4842
      _extenty        =   3334
      _stockprops     =   14
      caption         =   "Selección de Archivos..."
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
      Begin Threed.SSCheck NombreBase 
         Height          =   240
         Index           =   4
         Left            =   255
         TabIndex        =   7
         Top             =   1470
         Width           =   1905
         _version        =   65536
         _extentx        =   3360
         _extenty        =   423
         _stockprops     =   78
         caption         =   "  Base Hist_Vsd.MDB"
      End

      Begin Threed.SSCheck NombreBase 
         Height          =   240
         Index           =   3
         Left            =   255
         TabIndex        =   4
         Top             =   1200
         Width           =   1905
         _version        =   65536
         _extentx        =   3360
         _extenty        =   423
         _stockprops     =   78
         caption         =   "  Base Notarios.MDB"
      End

      Begin Threed.SSCheck NombreBase 
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   3
         Top             =   930
         Width           =   1905
         _version        =   65536
         _extentx        =   3360
         _extenty        =   423
         _stockprops     =   78
         caption         =   "  Base Oficina.MDB"
      End

      Begin Threed.SSCheck NombreBase 
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   2
         Top             =   660
         Width           =   1905
         _version        =   65536
         _extentx        =   3360
         _extenty        =   423
         _stockprops     =   78
         caption         =   "  Base EnvioSRM.MDB"
      End

      Begin Threed.SSCheck NombreBase 
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   1
         Top             =   390
         Width           =   1905
         _version        =   65536
         _extentx        =   3360
         _extenty        =   423
         _stockprops     =   78
         caption         =   "  Base Visado.MDB"
      End
   End

   Begin Threed.SSCommand Cancelar 
      Height          =   615
      Left            =   1575
      TabIndex        =   0
      Top             =   2025
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      font3d          =   3
      picture         =   "REPARABD.frx":0000
   End

   Begin Threed.SSCommand Aceptar 
      Height          =   615
      Left            =   450
      TabIndex        =   5
      Top             =   2025
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Aceptar"
      forecolor       =   8421376
      font3d          =   3
      picture         =   "REPARABD.frx":031A
   End
End
Attribute VB_Name = "RepararArchivos"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Aceptar_Click()
Dim I As Integer
Dim NombreArchivoOriginal As String
Dim NombreArchivoTemporal As String
Dim ContadorBasesParaReparar As Integer
Dim ContadorBasesMalReparadas As Integer
  If HaySeleccionados Then
    I = MsgBox("¿ Seguro de Reparar ?", 4 + 32 + 256, "Confirmar Reparación")
    If I = 6 Then
      Screen.MousePointer = 11
      ContadorBasesParaReparar = 0
      ContadorBasesMalReparadas = 0
      BaseVisado.Close
  On Error GoTo FalloReparacion
  
      For I = 0 To 4
        If NombreBase(I).Value Then
          ContadorBasesParaReparar = ContadorBasesParaReparar + 1
          NombreArchivoOriginal = DIRECTORIO_01 + Trim$(Mid(NombreBase(I).Caption, 8))
          NombreArchivoTemporal = DIRECTORIO_01 + Mid$(ARCHIVO_MDB05, 1, 6) + LTrim$(Str(I)) + _
                                                       Mid$(ARCHIVO_MDB05, 8)
          If Not ExisteArchivoCompactado(NombreArchivoOriginal, NombreArchivoTemporal, _
                                                               ContadorBasesMalReparadas) Then
            DBEngine.RepairDatabase NombreArchivoOriginal
            DBEngine.CompactDatabase NombreArchivoOriginal, NombreArchivoTemporal
            Call BorrarArchivo(NombreArchivoOriginal)
            Call RenombrarArchivo(NombreArchivoOriginal, NombreArchivoTemporal)
          End If
          NombreBase(I).Value = False
          NombreBase(I).Enabled = False
        End If
      Next I
      Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
      
  On Error GoTo 0
      Screen.MousePointer = 0
      Select Case ContadorBasesMalReparadas
        Case 0
          MsgBox "Los Archivos Fueron Reparados Exitosamente", 0 + 64, "Estado Reparación"
        Case ContadorBasesParaReparar
          MsgBox "Ningún Archivo Pudo Ser Reparado", 0 + 64, "Estado Reparación"
        Case Else
          MsgBox "No Todos los Archivos Fueron Reparados", 0 + 64, "Estado Reparación"
      End Select
      Cancelar.SetFocus
    End If
  Else
    MsgBox "No Hay Archivo Seleccionado", 0 + 16, "Faltan Parámetros"
  End If
  Exit Sub
  
FalloReparacion:
  ContadorBasesMalReparadas = ContadorBasesMalReparadas + 1
  MsgBox "No se Pudo Reparar Base" + Chr(10) + NombreArchivoOriginal, 0 + 48, "Archivo Irreparable"
  Resume Next
End Sub
Private Sub Aceptar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub
Private Sub Cancelar_Click()
  Unload Me
End Sub

Private Function HaySeleccionados() As Boolean
Dim I As Integer
  HaySeleccionados = False
  For I = 0 To 4
    If NombreBase(I).Value Then
      HaySeleccionados = True
      Exit For
    End If
  Next I
End Function
Private Function ExisteArchivoCompactado(ByVal ArchivoOriginal, ByVal ArchivoTemporal As String, Contador As Integer) As Boolean
Dim I As Integer
  ExisteArchivoCompactado = False
  If Dir(ArchivoTemporal) = UCase$(Mid(ArchivoTemporal, Len(ArchivoTemporal) - 10)) Then
    ExisteArchivoCompactado = True
    I = MsgBox("Archivo " + Chr(34) + ArchivoOriginal + Chr(34) + Chr(10) + "Ya Se Intentó Reparar Anteriormente", _
                        5 + 48 + 0, "Reparación Incompleta")
    If I = 4 Then
      Call BorrarArchivo(ArchivoTemporal)
      ExisteArchivoCompactado = False
    Else
      Contador = Contador + 1
    End If
  End If
End Function
Private Sub BorrarArchivo(ByVal ArchivoOriginal As String)
  Kill ArchivoOriginal
  ArchivoOriginal = Mid$(ArchivoOriginal, 1, Len(ArchivoOriginal) - 3) + "LDB"
  Kill ArchivoOriginal
End Sub
Private Sub RenombrarArchivo(ByVal ArchivoOriginal, ByVal ArchivoTemporal)
  Name ArchivoTemporal As ArchivoOriginal
  ArchivoOriginal = Mid$(ArchivoOriginal, 1, Len(ArchivoOriginal) - 3) + "LDB"
  ArchivoTemporal = Mid$(ArchivoTemporal, 1, Len(ArchivoTemporal) - 3) + "LDB"
  Name ArchivoTemporal As ArchivoOriginal
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub
Private Sub Form_Activate()
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
Private Sub NombreBase_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Aceptar_Click
    Case 27
      Unload Me
  End Select
End Sub
