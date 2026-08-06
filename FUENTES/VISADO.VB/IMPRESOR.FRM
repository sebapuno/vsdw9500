VERSION 4.00
Begin VB.Form DefineImpresora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo Impresora"
   ClientHeight    =   1635
   ClientLeft      =   4005
   ClientTop       =   3345
   ClientWidth     =   3075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   2040
   Left            =   3945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   Top             =   3000
   Width           =   3195
   Begin Threed.SSFrame SSFrame1 
      Height          =   960
      Left            =   45
      TabIndex        =   4
      Top             =   -45
      Width           =   2985
      _version        =   65536
      _extentx        =   5265
      _extenty        =   1693
      _stockprops     =   14
      caption         =   "Selección Impresora a Usar ..."
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
      Begin Threed.SSOption OpcionImpresora 
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   630
         Width           =   2450
         _version        =   65536
         _extentx        =   4313
         _extenty        =   423
         _stockprops     =   78
         caption         =   "&HP modelos DeskJet "
      End

      Begin Threed.SSOption OpcionImpresora 
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   2130
         _version        =   65536
         _extentx        =   3757
         _extenty        =   423
         _stockprops     =   78
         caption         =   "&Epson modelo FX-1170"
      End
   End

   Begin Threed.SSCommand Aceptar 
      Height          =   615
      Left            =   1215
      TabIndex        =   3
      Top             =   990
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Aceptar"
      forecolor       =   8421376
      font3d          =   3
      picture         =   "IMPRESOR.frx":0000
   End

   Begin Threed.SSCommand Cancelar 
      Height          =   615
      Left            =   2115
      TabIndex        =   0
      Top             =   990
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1085
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      font3d          =   3
      picture         =   "IMPRESOR.frx":031A
   End
End
Attribute VB_Name = "DefineImpresora"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Aceptar_Click()
Dim ImpresoraTemporal As String
  ImpresoraTemporal = DefinicionImpresora
  If OpcionImpresora(0).Value Then
    DefinicionImpresora = "ibm"
  Else
    DefinicionImpresora = "desk"
  End If
  If Not PoneValorINI("Visado", "Impresora", DefinicionImpresora, DIRECTORIO_01 + ARCHIVO_INI01) Then
    DefinicionImpresora = ImpresoraTemporal
    MsgBox "Imposible Actualizar Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 16, "Error de Actualización"
  End If
  Unload Me
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
  Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
  Call CentrarVentana(Me)
  If DefinicionImpresora = "ibm" Then
    OpcionImpresora(0).Value = True
  Else
    OpcionImpresora(1).Value = True
  End If
End Sub
Private Sub OpcionImpresora_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call Aceptar_Click
    Case 27
      Unload Me
  End Select
End Sub
