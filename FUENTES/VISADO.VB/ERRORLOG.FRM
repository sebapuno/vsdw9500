VERSION 4.00
Begin VB.Form ErrorLogueo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Problemas de Logueo"
   ClientHeight    =   2625
   ClientLeft      =   1575
   ClientTop       =   3315
   ClientWidth     =   7320
   Height          =   3030
   Left            =   1515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Top             =   2970
   Width           =   7440
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   1545
      Width           =   1380
      _version        =   65536
      _extentx        =   2434
      _extenty        =   450
      _stockprops     =   15
      caption         =   "Glosa Error"
      forecolor       =   0
      backcolor       =   -2147483644
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelouter      =   0
      font3d          =   2
      alignment       =   1
      autosize        =   2
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   915
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   90
      Width           =   6915
      _version        =   65536
      _extentx        =   12197
      _extenty        =   1614
      _stockprops     =   15
      caption         =   " Atención :"
      forecolor       =   192
      backcolor       =   -2147483644
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelinner      =   1
      font3d          =   1
      alignment       =   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"ERRORLOG.frx":0000
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   450
         Left            =   195
         TabIndex        =   4
         Top             =   300
         Width           =   6585
         WordWrap        =   -1  'True
      End
   End
   Begin Threed.SSPanel Glosa 
      Height          =   300
      Left            =   1710
      TabIndex        =   2
      Top             =   1545
      Width           =   5505
      _version        =   65536
      _extentx        =   9710
      _extenty        =   529
      _stockprops     =   15
      forecolor       =   128
      backcolor       =   -2147483644
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
      bevelinner      =   1
      font3d          =   3
      alignment       =   1
   End
   Begin Threed.SSPanel Tipo 
      Height          =   300
      Left            =   1710
      TabIndex        =   1
      Top             =   1125
      Width           =   2460
      _version        =   65536
      _extentx        =   4339
      _extenty        =   529
      _stockprops     =   15
      forecolor       =   128
      backcolor       =   -2147483644
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
      bevelinner      =   1
      font3d          =   3
      alignment       =   1
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   0
      Top             =   1125
      Width           =   1440
      _version        =   65536
      _extentx        =   2540
      _extenty        =   450
      _stockprops     =   15
      caption         =   "Tipo de Logueo"
      forecolor       =   0
      backcolor       =   -2147483644
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderwidth     =   1
      bevelouter      =   0
      font3d          =   2
      alignment       =   1
      autosize        =   1
   End
   Begin Threed.SSCommand Imprimir 
      Height          =   645
      Left            =   5490
      TabIndex        =   7
      Top             =   1935
      Width           =   780
      _version        =   65536
      _extentx        =   1376
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Imprimir"
      forecolor       =   8388608
      font3d          =   3
      autosize        =   1
      picture         =   "ERRORLOG.frx":0060
   End
   Begin Threed.SSCommand Cancelar 
      Height          =   645
      Left            =   6300
      TabIndex        =   6
      Top             =   1935
      Width           =   780
      _version        =   65536
      _extentx        =   1376
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      font3d          =   3
      picture         =   "ERRORLOG.frx":0832
   End
End
Attribute VB_Name = "ErrorLogueo"
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

Private Sub Form_Activate()
  Form.MousePointer = 0
  Cancelar.SetFocus
End Sub
Private Sub Form_Load()
  Call CentrarVentana(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Form.MousePointer = 11
End Sub

Private Sub Imprimir_Click()
  Screen.MousePointer = 11
On Error GoTo ErrorDeImpresion
  Printer.Print "Banco de Chile"; Tab(54); "Fecha Emisión : " + Format$(Date, "dd/mmm/yyyy")
  Printer.Print Tab(54); "Hora Emisión  : " + Format$(Time, "hh:mm:ss")
  Printer.Print
  Printer.FontUnderline = True
  Printer.Print Tab(58); "DETALLE ERROR DE LOGUEO"
  Printer.FontUnderline = False
  Printer.Print
  Printer.Print Tab(7); "Tipo de Logueo : " + Tipo.Caption
  Printer.Print Tab(7); "Glosa Error : " + Glosa.Caption
  Printer.EndDoc
On Error GoTo 0
  GoTo Fin
  
ErrorDeImpresion:
   I% = MsgBox("Error de Impresión", 5 + 48 + 0, "Impresión Cancelada")
   If I% = 4 Then
    Resume 0
   Else
    Printer.KillDoc
On Error GoTo 0
   End If
   
Fin:
  Screen.MousePointer = 0
  Cancelar.SetFocus
End Sub
