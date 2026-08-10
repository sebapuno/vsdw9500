VERSION 4.00
Begin VB.Form ErrorComunicacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Problemas de Comunicación"
   ClientHeight    =   4395
   ClientLeft      =   2070
   ClientTop       =   2625
   ClientWidth     =   7320
   Height          =   4740
   Left            =   2040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Top             =   2310
   Width           =   7380
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel NombreAplicacion 
      Height          =   300
      Left            =   1815
      TabIndex        =   19
      Top             =   3330
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   3330
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Left            =   90
      TabIndex        =   15
      Top             =   1725
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   1140
      Left            =   195
      TabIndex        =   13
      Top             =   90
      Width           =   1
   End
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Llave 
      Height          =   300
      Left            =   1815
      TabIndex        =   12
      Top             =   1725
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   2115
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel RetFuncion 
      Height          =   300
      Left            =   1815
      TabIndex        =   10
      Top             =   2115
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   2505
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Status 
      Height          =   300
      Left            =   1815
      TabIndex        =   8
      Top             =   2505
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   2925
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel ServidorDatos 
      Height          =   300
      Left            =   1815
      TabIndex        =   6
      Top             =   2925
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Nodo 
      Height          =   300
      Left            =   1500
      TabIndex        =   5
      Top             =   1305
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   1305
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Servidor 
      Height          =   300
      Left            =   3180
      TabIndex        =   3
      Top             =   1305
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   300
      Left            =   2190
      TabIndex        =   2
      Top             =   1305
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Operacion 
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   1305
      Width           =   1
   End
   ' FI2026 - VBX.SSPanel migrado a Threed.SSPanel (THREED32.OCX)
   Begin Threed.SSPanel Panel3D1 
      Height          =   300
      Left            =   3870
      TabIndex        =   0
      Top             =   1305
      Width           =   1
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

