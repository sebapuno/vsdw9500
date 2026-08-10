VERSION 4.00
Begin VB.Form ReferenciasVisado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Referencias"
   ClientHeight    =   1125
   ClientLeft      =   1770
   ClientTop       =   2055
   ClientWidth     =   5970
   ClipControls    =   0   'False
   Height          =   1530
   KeyPreview      =   -1  'True
   Left            =   1710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Top             =   1710
   Width           =   6090
   Begin Threed.SSPanel Label2 
      Height          =   375
      Left            =   1050
      TabIndex        =   0
      Top             =   150
      Width           =   4740
      _version        =   65536
      _extentx        =   8361
      _extenty        =   661
      _stockprops     =   15
      caption         =   "Sistema  :  Visado de Productos Windows"
      forecolor       =   8388608
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Times New Roman"
         charset         =   1
         weight          =   700
         size            =   10.5
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelouter      =   0
      font3d          =   3
      alignment       =   1
   End

   Begin Threed.SSPanel Label1 
      Height          =   375
      Left            =   1050
      TabIndex        =   1
      Top             =   600
      Width           =   2130
      _version        =   65536
      _extentx        =   3757
      _extenty        =   661
      _stockprops     =   15
      caption         =   "Versión  :  "
      forecolor       =   8388608
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Times New Roman"
         charset         =   1
         weight          =   700
         size            =   10.5
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelouter      =   0
      font3d          =   3
      alignment       =   1
   End

   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Picture         =   "REFERENC.frx":0000
      Top             =   150
      Width           =   705
   End
End
Attribute VB_Name = "ReferenciasVisado"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
  If (KeyAscii = 13) Or (KeyAscii = 27) Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  Principal.Enabled = False
  Call CentrarVentana(Me)
  'Version = SacaValorINI("Visado", "Version", DIRECTORIO_01 + ARCHIVO_INI01)
  Label1.Caption = "Versión  :  " + Version '2.0'"Sistema Visado " + Version + " Para Windows"
  
  Label1.FontSize = 12
  Label2.FontSize = 12
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Principal.Enabled = True
End Sub

