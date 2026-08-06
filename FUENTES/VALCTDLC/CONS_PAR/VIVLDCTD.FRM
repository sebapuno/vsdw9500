VERSION 4.00
Begin VB.Form VisorValidaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprimir Validaciones Cuenta Corriente y Línea de Crédito"
   ClientHeight    =   5370
   ClientLeft      =   1125
   ClientTop       =   1710
   ClientWidth     =   9435
   ControlBox      =   0   'False
   ForeColor       =   &H00800000&
   Height          =   5775
   Left            =   1065
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Top             =   1365
   Width           =   9555
   Begin Threed.SSFrame SSFrame1 
      Height          =   5295
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   9330
      _version        =   65536
      _extentx        =   16457
      _extenty        =   9340
      _stockprops     =   14
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
      ' FI2026 - VBX.SpreadSheet migrado a MSFlexGrid (MSFLXGRD.OCX)
      Begin MSFlexGridLib.MSFlexGrid Grilla 
         Height          =   3540
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Width           =   9105
         Cols            =   2
         Rows            =   2
         ScrollBars      =   2  'flexScrollBarVertical
         SelectionMode   =   1  'flexSelectionByRow
      End

      Begin Threed.SSFrame Frame3D1 
         Height          =   990
         Left            =   150
         TabIndex        =   8
         Top             =   4125
         Width           =   6255
         _version        =   65536
         _extentx        =   11033
         _extenty        =   1746
         _stockprops     =   14
         caption         =   "Selección a Desplegar"
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         font3d          =   3
         Begin VB.ComboBox cb_Orig 
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "VIVLDCTD.frx":0004
            Left            =   3825
            List            =   "VIVLDCTD.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   525
            Width           =   1140
         End

         Begin VB.ComboBox cb_Oper 
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "VIVLDCTD.frx":0021
            Left            =   2175
            List            =   "VIVLDCTD.frx":0031
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   525
            Width           =   1515
         End

         Begin VB.ComboBox cb_Prod 
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "VIVLDCTD.frx":0061
            Left            =   150
            List            =   "VIVLDCTD.frx":006B
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   525
            Width           =   1890
         End

         Begin Threed.SSPanel SSPanel1 
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   300
            Width           =   1890
            _version        =   65536
            _extentx        =   3334
            _extenty        =   423
            _stockprops     =   15
            caption         =   "Producto"
            forecolor       =   -2147483630
            backcolor       =   13160660
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   -1  'True
               strikethrough   =   0   'False
            EndProperty
            bevelouter      =   0
            font3d          =   3
            alignment       =   1
         End

         Begin Threed.SSPanel SSPanel1 
            Height          =   240
            Index           =   1
            Left            =   2175
            TabIndex        =   16
            Top             =   300
            Width           =   1290
            _version        =   65536
            _extentx        =   2275
            _extenty        =   423
            _stockprops     =   15
            caption         =   "Operación"
            forecolor       =   -2147483630
            backcolor       =   13160660
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   -1  'True
               strikethrough   =   0   'False
            EndProperty
            bevelouter      =   0
            font3d          =   3
            alignment       =   1
         End

         Begin Threed.SSPanel SSPanel1 
            Height          =   240
            Index           =   2
            Left            =   3825
            TabIndex        =   17
            Top             =   300
            Width           =   1065
            _version        =   65536
            _extentx        =   1879
            _extenty        =   423
            _stockprops     =   15
            caption         =   "Origen"
            forecolor       =   -2147483630
            backcolor       =   13160660
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   400
               size            =   8.25
               underline       =   0   'False
               italic          =   -1  'True
               strikethrough   =   0   'False
            EndProperty
            bevelouter      =   0
            font3d          =   3
            alignment       =   1
         End

         Begin Threed.SSCommand btn_Consultar 
            Height          =   645
            Left            =   5250
            TabIndex        =   18
            Top             =   225
            Width           =   780
            _version        =   65536
            _extentx        =   1376
            _extenty        =   1138
            _stockprops     =   78
            caption         =   "&Consultar"
            forecolor       =   16384
            font3d          =   3
            picture         =   "VIVLDCTD.frx":0093
         End
      End

      Begin Threed.SSCommand btn_Actualiza 
         Height          =   645
         Left            =   6600
         TabIndex        =   3
         Top             =   4350
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Actualiza"
         forecolor       =   8421376
         font3d          =   3
         picture         =   "VIVLDCTD.frx":051D
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Categoría"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   225
         Width           =   795
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Tipo Dato"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   6
         Left            =   8025
         TabIndex        =   13
         Top             =   225
         Width           =   945
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Largo"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   7155
         TabIndex        =   12
         Top             =   225
         Width           =   855
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Posición"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   6280
         TabIndex        =   11
         Top             =   225
         Width           =   860
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Argumento"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   930
         TabIndex        =   10
         Top             =   225
         Width           =   840
      End

      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Descripción"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   1785
         TabIndex        =   9
         Top             =   225
         Width           =   4480
      End

      Begin Threed.SSCommand Cancelar 
         Height          =   645
         Left            =   8400
         TabIndex        =   5
         Top             =   4350
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   1
         picture         =   "VIVLDCTD.frx":0837
      End

      Begin Threed.SSCommand Imprimir 
         Height          =   645
         Left            =   7500
         TabIndex        =   4
         Top             =   4350
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Imprimir"
         font3d          =   1
         autosize        =   1
         picture         =   "VIVLDCTD.frx":1049
      End
   End
End
Attribute VB_Name = "VisorValidaciones"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub btn_Actualiza_Click()

On Error GoTo Final

Screen.MousePointer = 11

'Me.Enabled = False
'Load Frm_BajCTD

'Unload Frm_BajCTD
'Me.Enabled = True

Set BaseEnvioSRM = OpenDatabase("C:\DATA\VISADO\ENVIOSRM.MDB")
Set BaseVldCta = OpenDatabase("C:\DATA\VISADO\VALCTDLC.MDB")

IdUsuario = ObtieneRutUsuarioLogueado()

BaseVldCta.Execute "Delete * From tbl_ValMET"
BaseVldCta.Execute "Delete * From tbl_ValREG"

Terminar = False

'IncrementoStatus = Int(3570 / 64)
'Status.Width = 15

'PanelStatus.Visible = True
'Status.Visible = True

Bajar_Datos ("4086")
Bajar_Datos ("4087")

'Status.Width = 3570

BaseEnvioSRM.Close
BaseVldCta.Close


MsgBox "Tablas Actualizadas en forma exitosa.", 64, "Actualización Tablas de Parámetros"

Screen.MousePointer = 0

Exit Sub
 
Final:
  Me.Enabled = True
  Screen.MousePointer = 0

End Sub


Private Sub btn_Consultar_Click()

Dim BaseVLDCTD As Database
Dim PuntVld As Recordset

Screen.MousePointer = 11

Grilla.Rows = 0

Tabla = IIf(UCase(cb_Orig.Text) = "METRO", "Tbl_ValMet", "Tbl_ValReg")
TipoProd = IIf(UCase(cb_Prod.Text) = "CUENTA CORRIENTE", "0", "1")

Select Case UCase(cb_Oper.Text)
     Case "APERTURA"
        CodOper = "0" 'OPER_APERTURA   ' 0
     Case "CIERRE"
        CodOper = "1" 'OPER_CIERRE   ' 1
     Case "RENOVACIÓN"
        CodOper = "2" 'OPER_RENOVACION   ' 2
     Case "MODIFICACIÓN"
        CodOper = "3" 'OPER_MODIFICACION   ' 3
  End Select

Set BaseVLDCTD = OpenDatabase("C:\DATA\VISADO\VALCTDLC.MDB")

Set PuntVld = BaseVLDCTD.OpenRecordset("Select * From " + Tabla + " Where CTD_Prod=" + TipoProd + " And CTD_Oper=" + CodOper + " Order by CTD_Categ,CTD_Argum Asc", dbOpenDynaset)
    '  CTD_Orig,CTD_Categ,CTD_PosVista
If PuntVld.RecordCount > 0 Then
  PuntVld.MoveLast
  PuntVld.MoveFirst
   
  For j% = 1 To PuntVld.RecordCount
    Grilla.Rows = Grilla.Rows + 1
    Grilla.Row = Grilla.Rows
    Grilla.Col = 1
    Grilla.Text = PuntVld!Ctd_Categ
    Grilla.Col = 2
    Grilla.Text = PuntVld!Ctd_Argum
    Grilla.Col = 3
    Grilla.Text = Trim$(PuntVld!Ctd_Descrip)
    Grilla.Col = 4
    Grilla.Text = PuntVld!Ctd_PosVista
    Grilla.Col = 5
    Grilla.Text = PuntVld!Ctd_LargoVta
    Grilla.Col = 6
    Grilla.Text = Mid$(PuntVld!Ctd_TipoDat, 1, 6)
    
    PuntVld.MoveNext
  Next j%
    
End If

PuntVld.Close
BaseVLDCTD.Close

Screen.MousePointer = 0

End Sub

Private Sub Cancelar_Click()
  Unload Me
  'GrillaProductosVisados.OperationMode = 0
   
End Sub

Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub




Private Sub cb_Oper_Click()

Grilla.Rows = 0

End Sub

Private Sub cb_Oper_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  cb_Orig.SetFocus
End If

End Sub


Private Sub cb_Orig_Click()

Grilla.Rows = 0

End Sub

Private Sub cb_Orig_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  btn_Actualiza.SetFocus
End If

End Sub


Private Sub cb_Prod_Click()

Grilla.Rows = 0

End Sub

Private Sub cb_Prod_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  cb_Oper.SetFocus
End If

End Sub


Private Sub Form_Load()
  
Dim Success As Integer
   
Screen.MousePointer = 11
  
If Dir("C:\DATA\VISADO\VALCTDLC.MDB") <> "VALCTDLC.MDB" Then
     MsgBox "No se encontró la base de datos VALCTDLC.MDB", 16, "Archivo no encontrado"
     Screen.MousePointer = 0
     Exit Sub
End If
  
Grilla.Cols = 6
Grilla.ColWidth(1) = 6.5
Grilla.ColWidth(2) = 7
Grilla.ColWidth(3) = 37.4
Grilla.ColWidth(4) = 7.1
Grilla.ColWidth(5) = 7.1
Grilla.ColWidth(6) = 7.6
  
cb_Prod.ListIndex = 0
cb_Oper.ListIndex = 0
cb_Orig.ListIndex = 0
 
Success = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZENOMOVE)
 
Screen.MousePointer = 0
  
End Sub


Private Sub Imprimir_Click()
Dim I As Integer
Dim Cadena As String
Dim TotPag As Integer
Dim PagActual As Integer
Dim FilaGrid As Integer

  Screen.MousePointer = 11
  
If Grilla.Rows > 0 Then
  
  If Grilla.Rows Mod 55 = 0 Then
     TotPag = Grilla.Rows \ 55
  Else
     TotPag = (Grilla.Rows \ 55) + 1
  End If
  FilaGrid = 0
  PagActual = 1
  
  Printer.FontName = "Courier New"
  Printer.FontBold = False
  Printer.FontSize = "12"
  Printer.Print "BANCO DE CHILE"; Tab(55); "Fecha : "; Format(Now, "dd/mm/yyyy");
  Printer.Print Tab(75); " 1/" + Trim(Str(TotPag))
  Printer.Print Tab(55); "Hora  : "; Format$(Time, "hh:mm:ss")
  Printer.Print
  Printer.FontUnderline = True
  Printer.FontBold = True
  Printer.Print Tab(6); "TABLA DE PARÁMETROS - VALIDACIÓN CUENTA CORRIENTE Y LÍNEA DE CRÉDITO "
  Printer.FontUnderline = False
  
  Printer.FontSize = "10"

  Printer.FontBold = False
  Printer.Print

  Printer.Print Tab(7); "Producto : ";
  
  Printer.FontBold = True
  Printer.Print UCase(cb_Prod.Text) + Space(4);
  Printer.FontBold = False
  
  Printer.Print "Operación : ";
  
  Printer.FontBold = True
  Printer.Print UCase(cb_Oper.Text) + Space(4);
  Printer.FontBold = False
  
  Printer.Print "Origen : ";
  
  Printer.FontBold = True
  Printer.Print UCase(cb_Orig.Text)
  Printer.FontBold = False
  
  
  Printer.Print
  Printer.FontSize = "10"
  Printer.Print Tab(7); String(87, "-")
  Printer.Print Tab(7); "CAT"; Tab(13); "ARGUM"; Tab(21); "DESCRIPCIÓN"; Tab(73); "POS"; Tab(78); "LARGO"; Tab(85); "TIPO DATO"
  Printer.Print Tab(7); String(87, "-")
  
  For I = 1 To Grilla.Rows
    FilaGrid = FilaGrid + 1
    Grilla.Row = I
    
    Grilla.Col = 1
    Cadena = Grilla.Text + Space(2)
    
    Grilla.Col = 2
    Cadena = Cadena + Grilla.Text + Space(2)
    
    Grilla.Col = 3
    Cadena = Cadena + Left$(Trim$(Grilla.Text) + Space(50), 50) + Space(2)
    
    Grilla.Col = 4 ' POS
    Cadena = Cadena + Right(Space(3) + Trim$(Str(Val(Grilla.Text))), 3) + Space(2)
    
    Grilla.Col = 5 ' LARGO
    Cadena = Cadena + Space(1) + Right(Space(3) + Trim$(Str(Val(Grilla.Text))), 3) + Space(2)
    
    Grilla.Col = 6 ' TIPO DATO
    Cadena = Cadena + Space(1) + Grilla.Text
    
    Printer.Print Tab(7); Cadena
    
    If FilaGrid = 55 And Grilla.Row < Grilla.Rows Then
        Printer.EndDoc
        PagActual = PagActual + 1
        FilaGrid = 0
        Printer.FontSize = "12"
        Printer.Print "BANCO DE CHILE"; Tab(55); "Fecha : "; Format(Now, "dd/mm/yyyy");
        Printer.Print Tab(75); Trim$(Str(PagActual)) + "/" + Trim(Str(TotPag))
        Printer.Print Tab(55); "Hora  : "; Format$(Time, "hh:mm:ss")
        Printer.Print
        'Printer.FontUnderline = True
        'Printer.FontBold = True
        'Printer.Print Tab(5); "TABLA DE PARÁMETROS - VALIDACIÓN CUENTA CORRIENTE Y LÍNEA DE CRÉDITO "
        'Printer.FontUnderline = False
        'Printer.FontBold = False
        'Printer.Print
        'Printer.Print Tab(5); "Producto : " + UCase(cb_Prod.Text) + Space(4) + "Operación : " + UCase(cb_Oper.Text) + Space(4) + "Origen : " + UCase(cb_Orig.Text)
        Printer.Print
        Printer.FontSize = "10"
        Printer.Print Tab(7); String(87, "-")
        Printer.Print Tab(7); "CAT"; Tab(13); "ARGUM"; Tab(21); "DESCRIPCIÓN"; Tab(73); "POS"; Tab(78); "LARGO"; Tab(85); "TIPO DATO"
        Printer.Print Tab(7); String(87, "-")
  
    End If
    
  Next I
  Printer.EndDoc
  
End If
  
Screen.MousePointer = 0
  
End Sub
Private Sub Imprimir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub













