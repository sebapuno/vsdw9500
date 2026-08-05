VERSION 4.00
Begin VB.Form ErrorComunicacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Problemas de Comunicación"
   ClientHeight    =   4395
   ClientLeft      =   2070
   ClientTop       =   2625
   ClientWidth     =   7320
   Height          =   4800
   Left            =   2010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Top             =   2280
   Width           =   7440
   Begin VBX.SSPanel NombreAplicacion 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1815
      TabIndex        =   19
      Top             =   3330
      Width           =   5400
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      AutoSize        =   3  'AutoSize Child To Panel
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Aplicación"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   18
      Top             =   3330
      Width           =   1635
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      AutoSize        =   2  'AutoSize Panel Height To Caption
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "LLave Acceso"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   1725
      Width           =   1380
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   0  'Left Justify - TOP
      BevelInner      =   1  'Inset
      BorderWidth     =   1
      Caption         =   " Atención :"
      Font3D          =   1  'Raised w/light shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H000000C0&
      Height          =   1140
      Index           =   1
      Left            =   195
      TabIndex        =   13
      Top             =   90
      Width           =   6915
      Begin VB.Label Mensaje 
         Alignment       =   2  'Center
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
         Height          =   765
         Left            =   200
         TabIndex        =   14
         Top             =   300
         Width           =   6550
         WordWrap        =   -1  'True
      End
   End
   Begin VBX.SSPanel Llave 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1815
      TabIndex        =   12
      Top             =   1725
      Width           =   5400
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   0  'Left Justify - TOP
      AutoSize        =   3  'AutoSize Child To Panel
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Retorno SRM"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   2115
      Width           =   1470
   End
   Begin VBX.SSPanel RetFuncion 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1815
      TabIndex        =   10
      Top             =   2115
      Width           =   5400
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      AutoSize        =   2  'AutoSize Panel Height To Caption
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Status"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   90
      TabIndex        =   9
      Top             =   2505
      Width           =   915
   End
   Begin VBX.SSPanel Status 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1815
      TabIndex        =   8
      Top             =   2505
      Width           =   5400
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      AutoSize        =   3  'AutoSize Child To Panel
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Servidor de Datos"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   90
      TabIndex        =   7
      Top             =   2925
      Width           =   1635
   End
   Begin VBX.SSPanel ServidorDatos 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1815
      TabIndex        =   6
      Top             =   2925
      Width           =   5400
   End
   Begin VBX.SSPanel Nodo 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   1500
      TabIndex        =   5
      Top             =   1305
      Width           =   615
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      AutoSize        =   3  'AutoSize Child To Panel
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Nodo Destino"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   1305
      Width           =   1335
   End
   Begin VBX.SSPanel Servidor 
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   3180
      TabIndex        =   3
      Top             =   1305
      Width           =   615
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Servidor"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   2190
      TabIndex        =   2
      Top             =   1305
      Width           =   915
   End
   Begin VBX.SSPanel Operacion 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelInner      =   1  'Inset
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Font3D          =   3  'Inset w/light shading
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   1305
      Width           =   2280
   End
   Begin VBX.SSPanel Panel3D1 
      Alignment       =   1  'Left Justify - MIDDLE
      BevelOuter      =   0  'None
      BorderWidth     =   1
      Caption         =   "Operación"
      Font3D          =   2  'Raised w/heavy shading
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   3870
      TabIndex        =   0
      Top             =   1305
      Width           =   1005
   End
   Begin Threed.SSCommand Imprimir 
      Height          =   645
      Left            =   5445
      TabIndex        =   17
      Top             =   3690
      Width           =   780
      _version        =   65536
      _extentx        =   1376
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Imprimir"
      forecolor       =   8388608
      font3d          =   3
      autosize        =   1
      picture         =   "ERRORCOM.frx":0000
   End
   Begin Threed.SSCommand Cancelar 
      Height          =   645
      Left            =   6270
      TabIndex        =   16
      Top             =   3690
      Width           =   780
      _version        =   65536
      _extentx        =   1376
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      font3d          =   3
      picture         =   "ERRORCOM.frx":07D2
   End
End
Attribute VB_Name = "ErrorComunicacion"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Function FormateaRut(NumeroRut As String) As String
  If Len(NumeroRut) > 9 Then
    NumeroRut = Mid$(NumeroRut, 2)
  End If
  NumeroRut = String(9 - Len(NumeroRut), "0") + NumeroRut
  FormateaRut = Mid$(NumeroRut, 1, 2) + "." + Mid$(NumeroRut, 3, 3) + "." + _
                            Mid$(NumeroRut, 6, 3) + "-" + UCase$(Mid$(NumeroRut, 9, 1))
End Function
Private Sub Cancelar_Click()
  Unload Me
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub Form_Activate()
  Cancelar.SetFocus
End Sub
Private Sub Form_Load()
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
  Sucess% = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZENOMOVE)
End Sub

Private Sub Imprimir_Click()
  Screen.MousePointer = 11
On Error GoTo ErrorDeImpresion
  Printer.Print "Banco de Chile"; Tab(54); "Fecha Emisión : " + Format$(Date, "dd/mmm/yyyy")
  Printer.Print Tab(54); "Hora Emisión  : " + Format$(Time, "hh:mm:ss")
  Printer.Print
  Printer.FontUnderline = True
  Printer.Print Tab(55); "DETALLE ERROR DE COMUNICACION"
  Printer.FontUnderline = False
  Printer.Print
  Printer.FontBold = True
  Printer.Print Tab(5); "Rut Cliente : " + FormateaRut(Rut)
  Printer.Print Tab(5); "Producto : " + PRODUCTOACCESADO
  Printer.Print Tab(5); "Aplicación : " + Mid$(EJECUTABLE, 1, 5) + Producto + Mid$(EJECUTABLE, 9)
  Printer.FontBold = False
  Printer.Print
  Printer.Print Tab(7); "Nodo Destino : " + Nodo.Caption
  Printer.Print Tab(7); "Servidor : " + Servidor.Caption
  Printer.Print Tab(7); "Operación :" + Operacion.Caption
  Printer.Print Tab(7); "Llave de Acceso :" + Llave.Caption
  Printer.Print Tab(7); "Retorno SRM :" + RetFuncion.Caption
  Printer.Print Tab(7); "Status :" + Status.Caption
  Printer.Print Tab(7); "Servidor de Datos :" + ServidorDatos.Caption
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
Private Sub Imprimir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

