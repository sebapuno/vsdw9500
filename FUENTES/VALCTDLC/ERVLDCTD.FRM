VERSION 4.00
Begin VB.Form Frm_ERVL 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errores de Validación de Datos"
   ClientHeight    =   4800
   ClientLeft      =   2730
   ClientTop       =   2475
   ClientWidth     =   7335
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
   Height          =   5205
   Left            =   2670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Top             =   2130
   Width           =   7455
   Begin Threed.SSPanel Lb_Zona 
      Height          =   315
      Left            =   5850
      TabIndex        =   2
      Top             =   150
      Width           =   1290
      _version        =   65536
      _extentx        =   2275
      _extenty        =   556
      _stockprops     =   15
      caption         =   "Regional"
      backcolor       =   13160660
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
   End

   Begin Threed.SSPanel Lb_Prod 
      Height          =   315
      Left            =   1875
      TabIndex        =   3
      Top             =   150
      Width           =   3690
      _version        =   65536
      _extentx        =   6509
      _extenty        =   556
      _stockprops     =   15
      caption         =   "Línea de Crédito N° 00-000-00000-00"
      backcolor       =   13160660
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
   End

   Begin Threed.SSPanel lb_Accion 
      Height          =   315
      Left            =   225
      TabIndex        =   4
      Top             =   150
      Width           =   1365
      _version        =   65536
      _extentx        =   2408
      _extenty        =   556
      _stockprops     =   15
      caption         =   "Renovación"
      backcolor       =   13160660
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
   End

   Begin MSGrid.Grid GridErr 
      Height          =   3135
      Left            =   225
      TabIndex        =   5
      Top             =   675
      Width           =   6915
      _version        =   65536
      _extentx        =   12197
      _extenty        =   5530
      _stockprops     =   77
      backcolor       =   16777215
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      borderstyle     =   0
      cols            =   4
      fixedcols       =   0
      mouseicon       =   "ERVLDCTD.frx":0000
   End

   Begin Threed.SSCommand Cancelar 
      Height          =   645
      Left            =   6150
      TabIndex        =   1
      Top             =   3975
      Width           =   840
      _version        =   65536
      _extentx        =   1482
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Cancelar"
      forecolor       =   192
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      font3d          =   3
      picture         =   "ERVLDCTD.frx":001C
   End

   Begin Threed.SSCommand Imprimir 
      Height          =   645
      Left            =   5025
      TabIndex        =   0
      Top             =   3975
      Width           =   855
      _version        =   65536
      _extentx        =   1508
      _extenty        =   1138
      _stockprops     =   78
      caption         =   "&Imprimir"
      forecolor       =   8388608
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      font3d          =   3
      autosize        =   1
      picture         =   "ERVLDCTD.frx":082E
   End
End
Attribute VB_Name = "Frm_ERVL"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Cancelar_Click()


Unload Me

End Sub


Private Sub Form_Load()

Dim Success As Integer

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

PreparaGrid_Errores

If TipoProd = "0" Then
  Lb_Prod.Caption = "Cuenta Corriente"
ElseIf TipoProd = "1" Then
  Lb_Prod.Caption = "Línea de Crédito"
Else
  Lb_Prod.Caption = "Producto"
End If

Lb_Zona.Caption = MetReg

Select Case Operacion
   Case "O"
      lb_Accion.Caption = "Apertura"   ' 0
   Case "C"
      lb_Accion.Caption = "Cierre"   ' 1
   Case "R"
      lb_Accion.Caption = "Renovación"   ' 2
   Case "M"
      lb_Accion.Caption = "Modificación"  ' 3
End Select

CtdLcd = IIf(Len(CtdLcd) = 12, Mid$(CtdLcd, 2), CtdLcd)

If Origen = "BCH" Then
  Lb_Prod.Caption = Lb_Prod.Caption + " N° " + Format$(CtdLcd, "0-000-00000-00")
ElseIf Origen = "BAE" Then
  Lb_Prod.Caption = Lb_Prod.Caption + " N° " + Format$(CtdLcd, "0-00-00-000000")
Else
  Lb_Prod.Caption = Lb_Prod.Caption + " N° " + CtdLcd
End If

Success = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZENOMOVE)

i = 1
While Trim$(Errores(i)) <> "" And i < 100
  GridErr.Rows = GridErr.Rows + 1
  
  GridErr.Row = i
  
  GridErr.Col = 0
  GridErr.Text = Mid$(Errores(i), 1, 3)
  
  GridErr.Col = 1
  GridErr.Text = Trim$(Mid$(Errores(i), 4, 15))
  
  GridErr.Col = 2
  GridErr.Text = Trim$(Mid$(Errores(i), 19))
  
  i = i + 1
  
Wend

GridErr.RemoveItem GridErr.Rows - 1

If GridErr.Rows > 8 Then
  GridErr.ColWidth(2) = 4155
End If

If GridErr.Rows > 1 And GridErr.Row <> 0 Then
  GridErr.Row = 1
  GridErr.Col = GridErr.Cols - 1
  Ilumina_Row GridErr
End If
GridErr.Refresh

End Sub


Public Sub PreparaGrid_Errores()

GridErr.Row = 0

GridErr.Col = 0
GridErr.Text = "Posición"
GridErr.FixedAlignment(0) = 2
GridErr.Col = 1
GridErr.Text = "Dato"
GridErr.FixedAlignment(1) = 2
GridErr.Col = 2
GridErr.Text = "Descripción"
GridErr.FixedAlignment(2) = 2

GridErr.ColWidth(0) = 1000
GridErr.ColWidth(1) = 1450
GridErr.ColWidth(2) = 4395
GridErr.ColWidth(3) = 1

GridErr.ColAlignment(0) = 2  ' 0 izq , 1 der , 2 Cen
GridErr.ColAlignment(1) = 2

End Sub




Private Sub GridErr_Click()

If GridErr.Rows > 1 And GridErr.Row <> 0 Then
  GridErr.Col = GridErr.Cols - 1
  Ilumina_Row GridErr
End If

End Sub

Private Sub Imprimir_Click()

Screen.MousePointer = 11

On Error GoTo ErrorDeImpresion
  
  Printer.Print "Banco de Chile"; Tab(54); "Fecha Emisión : " + Format$(Date, "dd/mmm/yyyy")
  Printer.Print Tab(54); "Hora Emisión  : " + Format$(Time, "hh:mm:ss")
  Printer.Print
  Printer.FontUnderline = True
  Printer.FontName = "Courier New"
  Printer.FontSize = 10
  Printer.FontBold = True
  Printer.Print Tab(30); "DETALLE ERRORES DE VALIDACIÓN"
  Printer.FontUnderline = False
  Printer.FontBold = False
  Printer.Print
  Printer.Print
  Printer.Print Tab(10); "Producto"; Tab(20); ": " + Lb_Prod.Caption
  Printer.Print Tab(10); "Acción"; Tab(20); ": " + lb_Accion.Caption
  Printer.Print Tab(10); "Origen"; Tab(20); ": " + Lb_Zona.Caption
  Printer.Print
  Printer.Print
  Printer.Print Tab(10); "Pos." + Space(4) + "Dato" + Space(14) + "Descripción"
  Printer.Print Tab(10); String(60, "-")
  
  For s% = 1 To GridErr.Rows - 1
    GridErr.Row = s%
    
    GridErr.Col = 0
    Printer.Print Tab(10); GridErr.Text;
    GridErr.Col = 1
    Printer.Print Tab(18); Left$(GridErr.Text + Space(10), 10);
    GridErr.Col = 2
    Printer.Print Trim$(GridErr.Text)
    
  Next s%
  
  Printer.EndDoc

On Error GoTo 0
  GoTo Fin
  
ErrorDeImpresion:
   i% = MsgBox("Error de Impresión", 5 + 48 + 0, "Impresión Cancelada")
   If i% = 4 Then
     Resume 0
   Else
     Printer.KillDoc
On Error GoTo 0
   End If
   
Fin:
  Screen.MousePointer = 0
  Cancelar.SetFocus


End Sub





