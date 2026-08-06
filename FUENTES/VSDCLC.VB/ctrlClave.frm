VERSION 5.00
Begin VB.Form ctrclave 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Acceso"
   ClientHeight    =   4590
   ClientLeft      =   5865
   ClientTop       =   3390
   ClientWidth     =   5370
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4590
   ScaleWidth      =   5370
   Begin VB.CommandButton cancelar_cam 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   3950
      Width           =   1215
   End
   Begin VB.CommandButton aceptar_cam 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   3950
      Width           =   1215
   End
   Begin VB.Frame pnl_cambio 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   5055
      Begin VB.TextBox clv_new2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3465
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1185
         Width           =   795
      End
      Begin VB.TextBox clv_new1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3465
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   840
         Width           =   795
      End
      Begin VB.TextBox clv_ant 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3465
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   270
         Width           =   795
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   480
         ' FI2026 - Picture eliminado: ctrlClave.frx no disponible
         Stretch         =   -1  'True
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1800
         TabIndex        =   14
         Top             =   285
         Width           =   1650
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1800
         TabIndex        =   13
         Top             =   840
         Width           =   1530
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Repita Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1800
         TabIndex        =   12
         Top             =   1185
         Width           =   1485
      End
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Cambia 
      Caption         =   "Cambiar Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txt_clave 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3195
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox txt_rut 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3195
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   360
         ' FI2026 - Picture eliminado: ctrlClave.frx no disponible
         Stretch         =   -1  'True
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1710
         TabIndex        =   4
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rut Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1710
         TabIndex        =   3
         Top             =   315
         Width           =   1305
      End
   End
End
Attribute VB_Name = "ctrclave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FI2025 SHG --- SE AGREGA FORMULARIO A PROYECTO

Option Explicit

Dim cuenta_veces As Integer
Const CLAVE_INICIAL = "mesa"
Dim ADMINISTRADOR As Integer



Private Sub aceptar_cam_Click()
Dim aux_rut$, aux_pass$, aux_cl$, aux_newcl$, sqlq$
Dim r$, Li_Largo%

If Trim(txt_rut.text) <> "" And Trim(clv_ant.text) <> "" And Trim(clv_new1.text) <> "" And Trim(clv_new2.text) <> "" Then
  If LCase$(Trim(clv_new1.text)) <> LCase$(Trim(clv_new2.text)) Then
    Beep
    MsgBox "Las claves ingresadas como nuevas no son iguales.", 16, "Control de Acceso"
    clv_new1.text = ""
    clv_new2.text = ""
    clv_new1.SetFocus

  ElseIf Trim$(LCase$(clv_new1.text)) = CLAVE_INICIAL Or Trim$(LCase$(clv_new2.text)) = CLAVE_INICIAL Then
    Beep
    MsgBox "La nueva clave no puede ser igual a la clave de inicializaci" & Chr$(243) & "n.", 16, "Control de Acceso"
    clv_new1.text = ""
    clv_new2.text = ""
    clv_new1.SetFocus

  ElseIf LCase$(Trim(clv_ant.text)) = LCase$(Trim(clv_new1.text)) Then
    Beep
    MsgBox "La nueva clave no puede ser igual a la antigua.", 16, "Control de Acceso."
    clv_ant.text = ""
    clv_new1.text = ""
    clv_new2.text = ""
    clv_ant.SetFocus

  Else
    If Len(txt_rut.text) < 3 Then
      Beep
      MsgBox "Ingrese un rut v" & Chr$(225) & "lido.", 16, "Control de Acceso" 'jp2
      Exit Sub
    End If
    aux_rut$ = CStr(CLng(Left(txt_rut.text, Len(txt_rut.text) - 2)))
    aux_pass$ = ""
    aux_cl$ = ""
    aux_newcl$ = ""
    
   'FI2025 'jp2
    RUT_USUARIO = aux_rut$
    
'jp2    sqlq$ = "exec pro_smp0080 2 " & CStr(CLng(FUNCIONALIDAD)) & "," & CStr(CLng(RUT_USUARIO))
       sqlq$ = "select a.password "
       sqlq$ = sqlq$ & "from mcambio.dbo.tbl_operadores a, mcambio.dbo.tbl_seg_funcionalidad_producto b "
       sqlq$ = sqlq$ & "where "
       sqlq$ = sqlq$ & "b.rut_operador = " & aux_rut$ & " and "
       sqlq$ = sqlq$ & "b.rut_operador = a.rut_operador and "
       sqlq$ = sqlq$ & "a.no_vigente is null"

    ParamSrm8K.APartirDe = 1
    r$ = SendQuery(NAME_BD, SERV_BD, NODO_BD, sqlq$)
         
    
    If r$ <> "" Then
       Li_Largo% = 1
       aux_pass$ = Trim(Trae_Campo(r$, Li_Largo%))
       Li_Largo% = Li_Largo% + Len(aux_pass$) + 1

       aux_cl$ = encripta(LCase$(clv_ant.text))
   
       If aux_cl$ <> aux_pass$ Then
          Beep
          MsgBox "Clave anterior es incorrecta.", 16, "Control de Acceso"
          clv_ant.text = ""
          clv_ant.SetFocus
       Else
          aux_newcl$ = encripta(LCase$(clv_new1.text))
         'FI2025 TODO ver como actualizar pass mediante SP
          sqlq$ = "update tbl_operadores set password = '" & aux_newcl$
          sqlq$ = sqlq$ & "' where rut_operador = " & aux_rut$
          ParamSrm8K.APartirDe = 1
          r$ = SendQuery(NAME_BD, SERV_BD, NODO_BD, sqlq$)
          If r$ = "" Then
            MsgBox "La clave fue cambiada satisfactoriamente.", 64, "Control de Acceso"
            'pnl_cambio.Visible = False
            pnl_cambio.Enabled = False
            txt_rut.Visible = True
            txt_clave.Visible = True
            txt_clave.text = ""
            Me.Height = 2500
            Me.Width = 5460

            txt_clave.SetFocus
          Else
            Beep
            MsgBox "La clave no pudo ser cambiada debido a problemas durante el proceso de actualizaci" & Chr$(243) & "n.", 64, "Control de Acceso"
          End If
       End If

    Else
        Beep
        MsgBox "Usted no es un usuario v" & Chr$(225) & "lido.", 16, "Control de Acceso" 'jp2
        'pnl_cambio.Visible = False
        pnl_cambio.Enabled = False
        txt_clave.SetFocus
    End If
  End If
End If

End Sub
Private Sub Aceptar_Click()
Dim aux_rut$, aux_pass$, sqlq$, r$, Li_Largo%
Dim aux_dv$, aux_clini$, aux_cl$
Dim pr_funcion As Integer

''If IsNumeric(pr_funcion) < 1 Then
''    FUNCIONALIDAD = 88
''    pr_funcion = FUNCIONALIDAD
''End If

If Trim(txt_rut.text) <> "" And Trim(txt_clave.text) <> "" Then
    If Len(txt_rut.text) < 3 Then
       MsgBox "Ingrese un rut v" & Chr$(225) & "lido.", 16, "Control de Acceso" 'jp2
       Exit Sub
    End If
    aux_rut$ = CStr(CLng(Left(txt_rut.text, Len(txt_rut.text) - 2)))
    RUT_USUARIO = aux_rut$
    aux_pass$ = ""
      
    sqlq$ = "select a.password, dv_rut_operador "
    sqlq$ = sqlq$ & "from mcambio.dbo.tbl_operadores a, mcambio.dbo.tbl_seg_funcionalidad_producto b "
    sqlq$ = sqlq$ & "where b.id_funcion=" & FUNCIONALIDAD & " and "
    sqlq$ = sqlq$ & "b.rut_operador = " & aux_rut$ & " and "
    sqlq$ = sqlq$ & "b.rut_operador = a.rut_operador and "
    sqlq$ = sqlq$ & "a.no_vigente is null"
    
    '' sqlq$ = "select a.password, a.dv_rut_operador "
    ''   sqlq$ = sqlq$ & "from mcambio.dbo.tbl_operadores a "
    ''   sqlq$ = sqlq$ & "where "
    ''   sqlq$ = sqlq$ & "a.rut_operador = " & aux_rut$ & " and "
    ''   sqlq$ = sqlq$ & "a.no_vigente is null"
       
    'jp2 End If

    ParamSrm8K.APartirDe = 1
       
    r$ = SendQuery(NAME_BD, SERV_BD, NODO_BD, sqlq$)
    If r$ <> "" Then
       Li_Largo% = 1
       aux_pass$ = Trim(Trae_Campo(r$, Li_Largo%))
       Li_Largo% = Li_Largo% + Len(aux_pass$) + 1
       aux_dv$ = Trim(Trae_Campo(r$, Li_Largo%))

       aux_clini$ = encripta(CLAVE_INICIAL)
       aux_cl$ = encripta(LCase$(txt_clave.text))
   
       If aux_cl$ = aux_pass$ Then
          If aux_cl$ = aux_clini$ Then
             txt_clave.text = ""
             Beep
             MsgBox "Debe cambiar la clave de inicializaci?n por una propia.", 16, "Control de Acceso"
             cambia_Click
             Exit Sub
          End If
          RET_ACCESO = True
          RUT_USUARIO = aux_rut$
          DV_USUARIO = aux_dv$
          
' FI2025 SHG -- Se pasa informaci?? variables globales incorporadas

          xRUTusr = RUT_USUARIO
          xDVusr = DV_USUARIO
          
'          PASSWORD_SUPERVISOR = aux_pass$
          Unload Me
       Else
          Beep
          MsgBox "Clave incorrecta.", 16, "Control de Acceso"
          RET_ACCESO = False
          cuenta_veces = cuenta_veces + 1
          If cuenta_veces = 3 Then
             '------------------------------------------------------------------{
             ' Proy. Win7 - Estabilizaci?n para estaci?n W_7INTE1
             '  Realsystems - 07/2016 PAC
             ' Se agrega c?digo para que no continue al aplicativo si fall? login.
             '---------------------
             Beep
             MsgBox "Complet" & Chr$(243) & " 3 intentos fallidos. No tiene acceso al aplicativo.", 16, "Control de Acceso"
             RET_ACCESO = False
             '------------------------------------------------------------------}
             Unload Me
          Else
             txt_clave.text = ""
             txt_clave.SetFocus
          End If
       End If

    Else
        Beep
        MsgBox "Usted no est" & Chr$(225) & " autorizado para realizar esta acci" & Chr$(243) & "n.", 16, "Control de Acceso" 'jp2
        RET_ACCESO = False
        Unload Me
    End If

End If
End Sub

Private Sub cambia_Click()
If Trim(txt_rut.text) <> "" Then
   Me.Height = 4845
   Me.Width = 5460
   clv_ant.text = ""
   clv_new1.text = ""
   clv_new2.text = ""
   pnl_cambio.Enabled = True
   clv_ant.SetFocus
Else
   Beep
   MsgBox "Debe ingresar el Rut al cual est" & Chr$(225) & " asociada la clave.", 16, "Control de Acceso" 'jp2
End If
End Sub

Private Sub cancelar_cam_Click()
'pnl_cambio.Visible = False
pnl_cambio.Enabled = False
txt_clave.Visible = True
txt_rut.Visible = True
Me.Height = 2500
Me.Width = 5460

txt_clave.SetFocus
End Sub

Private Sub Cancelar_Click()
RET_ACCESO = False
Unload Me
End Sub

Private Sub clv_ant_Click()
clv_ant.SelStart = 0
clv_ant.SelLength = Len(clv_ant.text)
End Sub

Private Sub clv_ant_GotFocus()
clv_ant.SelStart = 0
clv_ant.SelLength = Len(clv_ant.text)
End Sub

Private Sub clv_ant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  KeyAscii = 0
 'jp2 SendKeys "{TAB}"
 clv_new1.SetFocus
ElseIf KeyAscii = 8 Then

ElseIf KeyAscii < 65 Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub clv_new1_Click()
clv_new1.SelStart = 0
clv_new1.SelLength = Len(clv_new1.text)
End Sub

Private Sub clv_new1_GotFocus()
clv_new1.SelStart = 0
clv_new1.SelLength = Len(clv_new1.text)
End Sub

Private Sub clv_new1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
'  SendKeys "{TAB}"
      clv_new2.SetFocus

ElseIf KeyAscii = 8 Then

ElseIf KeyAscii < 65 Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub clv_new1_LostFocus()
clv_new1.text = LCase$(clv_new1.text)
If Trim$(clv_new1.text) = CLAVE_INICIAL Then
   Beep
   MsgBox "La nueva clave no puede ser igual a la clave de inicializaci" & Chr$(243) & "n.", 16, "Control de Acceso"
   clv_new1.text = ""
   clv_new1.SetFocus
End If
End Sub

Private Sub clv_new2_Click()
clv_new2.SelStart = 0
clv_new2.SelLength = Len(clv_new2.text)
End Sub

Private Sub clv_new2_GotFocus()
clv_new2.SelStart = 0
clv_new2.SelLength = Len(clv_new2.text)
End Sub

Private Sub clv_new2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    aceptar_cam_Click

ElseIf KeyAscii = 8 Then

ElseIf KeyAscii < 65 Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub clv_new2_LostFocus()
clv_new2.text = LCase$(clv_new2.text)
If Trim$(clv_new2.text) = CLAVE_INICIAL Then
   Beep
   MsgBox "La nueva clave no puede ser igual a la clave de inicializaci" & Chr$(243) & "n.", 16, "Control de Acceso"
   clv_new2.text = ""
   clv_new2.SetFocus
End If

End Sub

Private Sub Form_Load()

  Dim flag As Integer
  
  flag = 1

On Error GoTo Error_Form_Load
    Me.Height = 2500
    Me.Width = 5460
    Me.Left = (Screen.Width \ 2) - (Me.Width \ 2)
    Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)
    RET_ACCESO = False
    'RUT_SUPERVISOR = ""
    'DV_SUPERVISOR = ""
  '  PASSWORD_SUPERVISOR = ""
    

    cuenta_veces = 0

    txt_clave.text = ""
    If flag = 0 Then
      txt_rut.text = Format(CLng(RUT_USUARIO), "#,##0") & "-" & DV_USUARIO
      txt_rut.Enabled = False
    ElseIf flag = 1 Then
      txt_rut.text = ""
    End If
    
Exit Sub

Error_Form_Load:
    MsgBox "Error " & CStr(Err) & ": " & Error$, vbCritical, "Control de Acceso"
    Exit Sub
End Sub

Private Sub txt_clave_Click()
txt_clave.SelStart = 0
txt_clave.SelLength = Len(txt_clave.text)
End Sub

Private Sub txt_clave_GotFocus()
txt_clave.SelStart = 0
txt_clave.SelLength = Len(txt_clave.text)
End Sub

Private Sub txt_clave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  KeyAscii = 0
  Aceptar_Click

ElseIf KeyAscii = 8 Then

ElseIf KeyAscii < 65 Or (KeyAscii > 90 And KeyAscii < 97) Or KeyAscii > 122 Then
  Beep
  KeyAscii = 0
End If

End Sub

Private Sub txt_rut_Click()
txt_rut.SelStart = 0
txt_rut.SelLength = Len(txt_rut.text)
End Sub

Private Sub txt_rut_GotFocus()
txt_rut.SelStart = 0
txt_rut.SelLength = Len(txt_rut.text)
End Sub

Private Sub txt_rut_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    'KeyAscii = 0
    txt_rut.text = ""
End If

End Sub

Private Sub txt_rut_KeyPress(KeyAscii As Integer)
Dim numrut As Long
Dim digito As String
Dim existe_K%, auxnum$, i%, a$

existe_K% = False
auxnum$ = ""
If Len(txt_rut.text) > 0 Then
    For i% = 1 To Len(txt_rut.text)
      a$ = Mid$(txt_rut.text, i%, 1)
      Select Case a$
      Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
           auxnum$ = auxnum$ & a$
      Case "K"
           auxnum$ = auxnum$ & a$
           existe_K% = True
      End Select
    Next
End If


 Select Case KeyAscii
    Case 48 To 57, 75, 107
      If Len(auxnum$) + 1 < 10 And Not existe_K% Then
         auxnum$ = auxnum$ & UCase$(Chr(KeyAscii))
      Else
         Beep
         KeyAscii = 0
      End If

   Case 13
      KeyAscii = 0
      txt_clave.SetFocus
      Exit Sub

Case 8
    Exit Sub
   Case Else
      Beep
      KeyAscii = 0
End Select

If Len(auxnum$) > 1 Then
   digito = Right(auxnum$, 1)
   numrut = CLng(Left(auxnum$, Len(auxnum$) - 1))
   txt_rut.text = Format(numrut, "#,##0") & "-" & digito
   KeyAscii = 0
   txt_rut.SelStart = Len(txt_rut.text)
End If

End Sub

Private Sub txt_rut_LostFocus()
Dim numrut As Long
Dim digito As String
Dim verdig As String

If Trim(txt_rut.text) <> "" And Len(txt_rut.text) > 2 Then
   digito = UCase(Right(txt_rut.text, 1))
   numrut = CLng(Left(txt_rut.text, Len(txt_rut.text) - 2))
   verdig = UCase$(DvCalc(CStr(numrut)))
   If UCase$(DvCalc(CStr(numrut))) <> digito Then
      Beep
      MsgBox "El d" & Chr$(237) & "gito verificador no corresponde.", 16, "Verificaci" & Chr$(243) & "n Digito Verificador" 'jp2
      txt_rut.text = ""
      txt_rut.SetFocus
   End If
Else
   txt_rut.text = ""
   If Trim(txt_rut.text) <> "" Then txt_rut.SetFocus
End If

End Sub

