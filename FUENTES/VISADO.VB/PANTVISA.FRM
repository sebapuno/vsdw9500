VERSION 4.00
Begin VB.Form PantallaVisar 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Pantalla de Visado"
   ClientHeight    =   5460
   ClientLeft      =   3525
   ClientTop       =   1680
   ClientWidth     =   9105
   Height          =   5865
   Left            =   3465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Top             =   1335
   Width           =   9225
   Begin VB.PictureBox Panel 
      BackColor       =   &H00404080&
      Height          =   3825
      Left            =   90
      ScaleHeight     =   3765
      ScaleWidth      =   4590
      TabIndex        =   19
      Top             =   1410
      Visible         =   0   'False
      Width           =   4650
      Begin VB.CommandButton BtnPauta 
         Height          =   203
         Index           =   0
         Left            =   510
         MouseIcon       =   "PANTVISA.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   192
      End

      Begin VB.PictureBox Aceptar 
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   3855
         ScaleHeight     =   285
         ScaleWidth      =   675
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End

      Begin VB.Label Numero 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   22
         Top             =   120
         Width           =   300
      End

      Begin VB.Label Glosa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa"
         BeginProperty Font 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   780
         TabIndex        =   21
         Top             =   120
         Width           =   3600
         WordWrap        =   -1  'True
      End
   End

   Begin VB.PictureBox Sombra 
      BackColor       =   &H00000000&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   3945
      Index           =   1
      Left            =   210
      ScaleHeight     =   3885
      ScaleWidth      =   4590
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   4650
   End

   Begin VB.PictureBox Panel3D3 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   750
      Left            =   7020
      ScaleHeight     =   690
      ScaleWidth      =   1815
      TabIndex        =   16
      Top             =   4095
      Width           =   1875
      Begin VB.PictureBox Imprimir 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   795
         TabIndex        =   2
         Top             =   60
         Width           =   855
      End

      Begin VB.PictureBox Cancelar 
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   945
         ScaleHeight     =   555
         ScaleWidth      =   795
         TabIndex        =   3
         Top             =   60
         Width           =   855
      End
   End

   Begin VB.PictureBox Panel3D1 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   750
      Left            =   5130
      ScaleHeight     =   690
      ScaleWidth      =   1800
      TabIndex        =   15
      Top             =   4095
      Width           =   1860
      Begin VB.PictureBox Pauta 
         ForeColor       =   &H00800080&
         Height          =   615
         Left            =   60
         ScaleHeight     =   555
         ScaleWidth      =   795
         TabIndex        =   0
         Top             =   60
         Width           =   855
      End

      Begin VB.PictureBox Aprobar 
         Enabled         =   0   'False
         ForeColor       =   &H00808000&
         Height          =   615
         Left            =   950
         ScaleHeight     =   555
         ScaleWidth      =   795
         TabIndex        =   1
         Top             =   60
         Width           =   855
      End
   End

   Begin VB.PictureBox Panel3D2 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   8790
      TabIndex        =   6
      Top             =   270
      Width           =   8850
      Begin VB.Label Producto 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4860
         TabIndex        =   14
         Top             =   540
         Width           =   3795
      End

      Begin VB.Label Accion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4860
         TabIndex        =   13
         Top             =   300
         Width           =   3795
      End

      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acción"
         BeginProperty Font 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   3
         Left            =   4050
         TabIndex        =   12
         Top             =   300
         Width           =   510
      End

      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
         BeginProperty Font 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   2
         Left            =   4050
         TabIndex        =   11
         Top             =   540
         Width           =   660
      End

      Begin VB.Label Nombre 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   900
         TabIndex        =   10
         Top             =   540
         Width           =   2535
      End

      Begin VB.Label Rut 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   900
         TabIndex        =   9
         Top             =   300
         Width           =   2535
      End

      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   540
         Width           =   585
      End

      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Rut"
         BeginProperty Font 
            name            =   "Times New Roman"
            charset         =   1
            weight          =   400
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   270
      End
   End

   Begin MSFlexGridLib.MSFlexGrid GrillaVisado
      Height          =   2325
      Left            =   120
      TabIndex        =   4
      Top             =   1410
      Width           =   8805
      Cols            =   5
      Rows            =   1
      ScrollBars      =   2  'flexScrollBarVertical
      SelectionMode   =   1  'flexSelectionByRow
   End

   Begin VB.Image Rechazado 
      Height          =   240
      Left            =   5715
      Picture         =   "PANTVISA.frx":030A
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End

   Begin VB.Image Aprobado 
      Height          =   240
      Left            =   5400
      Picture         =   "PANTVISA.frx":0A8C
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
   End

   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      Height          =   2400
      Left            =   180
      TabIndex        =   17
      Top             =   1485
      Width           =   8805
   End
End
Attribute VB_Name = "PantallaVisar"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Dim IndiceBotonParaVerFirma As Integer
Private Sub MuestraDetalleEnGrilla()
Dim Fila           As Integer
Dim Data           As String
Dim Llave          As String
Dim Linea          As String
Dim largo          As Integer
Dim Posicion       As Integer
Dim Columna        As Integer
Dim Producto       As String
Dim IndiceTabPrd   As Recordset
Dim FiltroTabDesp  As Recordset
Dim IndiceTabOper  As Recordset
Dim PosicionComa   As Integer
Dim NombreArchivo  As String

Dim ls_sql As String
Dim ls_convenio As String
Dim ls_rendicion As String
Dim ls_contin As String
Dim ls_ctacte As String
Dim ls_glosa As String
Dim inicio As Integer
Dim sigue As Boolean
Dim ls_ret As String
Dim li_film As Integer
Dim li_filo As Integer
'******************************************************************
'Desarrollado  : Entersoft
'Observaciones : Agrega nuevos campos solicitados grilla de detalle
'Fecha         : 10-09-2009
'******************************************************************
Dim ls_nom_convenio As String
Dim ls_vigencia As String
Dim ls_fecha_in As String
Dim ls_fecha_fin As String
Dim ls_ofi_rela As String
Dim ls_ofi_rendi As String
Dim ls_ofi_contrato As String
Dim ls_ctacte_interna As String
Dim ls_ctacte_empresa As String
Dim ls_ctacte_factu As String
Dim ls_act_economica As String
Dim ls_adm_rechazo As String
Dim ls_contrata_avb As String
'Dim ls_concilia As String  'campo rendicion
Dim ls_nom_contacto As String
Dim ls_fono_contacto As String
Dim ls_fax_contacto As String
Dim ls_visado As String
'*******************************************************************
Dim lc_convenio As String
Dim lc_rendicion As String

Dim ls_existe As String
Dim Li_Largo1 As Integer
Dim Li_Largo2 As Integer

Dim Usuario As String
'Dim HaydatoS As Boolean
Dim NodoDestino As String
Dim Desplazamiento As Integer
Dim AplicacionDestino As String * 4
  HaydatoS = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)



  Set IndiceTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  Set IndiceTabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  Visando.GrillaProductos.Col = 4
  Producto = Visando.GrillaProductos.Text
  Visando.GrillaProductos.Col = 6
  
  Data = Visando.GrillaProductos.Text
  Retorno% = PoneValorINI("Parametros", "LlaveVisad", Data, DIRECTORIO_01 + ARCHIVO_INI01)
  
  IndiceTabPrd.Index = "Tab_Prd"
  ' FI2026 - IndiceTabOper.Index = IndiceTabOper.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  IndiceTabPrd.Seek "=", Producto
  Llave = IndiceTabPrd("Prd_Pos_Act")
  If Interpreta(Llave, SIMBOLOPOSICION) <> -1 Then
    Llave = Mid$(Data, Interpreta(Llave, SIMBOLOPOSICION), Interpreta(Llave, SIMBOLOLARGO))
  End If
  IndiceTabOper.Seek "=", Producto, Llave
  If Not IndiceTabOper.NoMatch Then
    Accion.Caption = IndiceTabOper("Oper_Glosa")
  Else
    Accion.Caption = Llave
  End If
  IndiceTabPrd.Close
  IndiceTabOper.Close
  NombreArchivo = UCase(DIRECTORIO_02 + _
                               Mid$(ARCHIVO_EXE07, 1, Len(ARCHIVO_EXE07) - 7) + _
                               Producto + Mid$(ARCHIVO_EXE07, Len(ARCHIVO_EXE07) - 3, Len(ARCHIVO_EXE07)))
  If Dir(NombreArchivo) <> Mid$(NombreArchivo, Len(NombreArchivo) - 11) Then
    Set FiltroTabDesp = BaseVisado.OpenRecordset("SELECT * FROM Tab_Desp WHERE Desp_Prod='" + _
                                                                                       Producto + "';", dbOpenSnapshot)
                                                                                       
    While Not FiltroTabDesp.EOF And Not Recorte
      ' FI2026 - .Rows - 1 = filas de datos (era .MaxRows)
      If FiltroTabDesp("Desp_Fil") > GrillaVisado.Rows - 1 Then
        GrillaVisado.Rows = FiltroTabDesp("Desp_Fil") + 1 ' FI2026 - era .MaxRows
      End If
      GrillaVisado.Col = FiltroTabDesp("Desp_Col")
      GrillaVisado.Row = FiltroTabDesp("Desp_Fil")
      GrillaVisado.Text = FiltroTabDesp("Desp_Glosa")
      If FiltroTabDesp("Desp_Picture") <> "" Then
        GrillaVisado.Col = GrillaVisado.Col + 1
        If FiltroTabDesp("Desp_Pos") = "TAL" Then
          GrillaVisado.Text = FiltroTabDesp("Desp_Picture")
        Else
          PosicionComa = InStr(FiltroTabDesp("Desp_Pos"), ",")
          Posicion = Val(Mid(FiltroTabDesp("Desp_Pos"), 1, PosicionComa - 1))
          largo = Val(Mid(FiltroTabDesp("Desp_Pos"), PosicionComa + 1))
          GrillaVisado.Text = ObtenerDescripcionGlosa(Producto, Mid$(Data, Posicion, largo), FiltroTabDesp("Desp_Picture"))
        End If
      End If
      FiltroTabDesp.MoveNext
    Wend
    
      'If Producto = "SP0" Then
        GrillaVisado.Col = 2
        GrillaVisado.Row = 1
        gCodConvenio = GrillaVisado.Text
        
        GrillaVisado.Rows = GrillaVisado.Rows + 1
        li_filo = GrillaVisado.Rows - 1 ' FI2026 - era .MaxRows: cantidad de filas de DATOS

        inicio = 1
        sigue = True
        ''Query de Consulta por Medio de Pago y Operaciones
        ''*************************************************
        GRut_SForma = Visando.RutCliente.Caption
        ls_sql = ""
'*************************************************************************************************************************************************************************
'Modificado por        : Entersoft
'Fecha de modificacion : 13/08/2009
'Detalle               : Reeplazar en aplicacion Visado campo anx_con_tipo_conv de tabla anx_convenio por campo con_casilla de tabla convenio
'Fecha de modificacion : 21/08/2009
'Detalle               : Reemplazar en aplicacion visado campo anx_con_tipo_ren de tabla anx_convenio por campo con_entrega_vv de tabla convenio
        
'        ls_sql = ls_sql & "select anx_con_tipo_conv + '~' + anx_con_tipo_rend + '~' + anx_con_contingencia + '~' + isnull(anx_con_cta_rend, '0')  + '~' + 'convenio'+ '~'  "
'        ls_sql = ls_sql & "select con_casilla + '~' + con_entrega_vv + '~' + anx_con_contingencia + '~' + isnull(anx_con_cta_rend, '0')  + '~' + 'convenio'+ '~'  "
         ls_sql = ls_sql & "select con_casilla + '~' + con_entrega_vv  + '~' + anx_con_contingencia + '~' + isnull(anx_con_cta_rend, '0')  + '~' + 'convenio'+ '~' +"
         ls_sql = ls_sql & "con_nom_convenio + '~'  + con_vigencia  + '~' + "
         ls_sql = ls_sql & "convert(char(10),con_fecha_in,103)  + '~' + convert(char(10),con_fecha_fin,103) + '~' + "
         ls_sql = ls_sql & "con_ofi_rela  + '~' + con_ofi_rendi + '~' + "
         ls_sql = ls_sql & "con_ofi_contrato + '~'  + con_ctacte_interna + '~' + con_ctacte_empresa + '~' + con_ctacte_factu  + '~' + con_act_economica + '~' + con_adm_rechazo + '~' + "                          'campos nuevos
         ls_sql = ls_sql & "con_contrata_avb + '~'  + con_nom_contacto    + '~' + con_fono_contacto + '~' + con_fax_contacto  + '~' + con_visado + '~'  "                                                          'campos nuevos
'        ls_sql = ls_sql & "con_contrata_avb + '~'  + rendicion          + '~' + con_nom_contacto    + '~' + con_fono_contacto + '~' + con_fax_contacto  + '~' + con_visado + '~'  "
'*************************************************************************************************************************************************************************
        
        ls_sql = ls_sql & " from convenio, "
        ls_sql = ls_sql & "  anx_convenio "
        ls_sql = ls_sql & " where con_cliente = " & Mid(GRut_SForma, 1, Len(GRut_SForma) - 1) & " "
        ls_sql = ls_sql & " and con_convenio = '" & gCodConvenio & "'"
        ls_sql = ls_sql & " and con_cliente *= anx_con_rut_empresa "
        ls_sql = ls_sql & " and con_convenio *= anx_con_convenio "


        Do While sigue
            CadenaDatos = ""
            DoEvents
            CadenaDatos = Llamada_ConReintentar("U076", "GS23", "b12256781234R00000011npagos      ", ls_sql, "8K")
            ls_existe = Mid(CadenaDatos, 7, 1)
            
            If ls_existe = "N" Then sigue = False
            ls_ret = Mid(CadenaDatos, 15, Len(CadenaDatos))
            If ls_ret <> "" Then
                Li_Largo1 = Len(ls_ret)
                Li_Largo2 = 1
                
                Do While Li_Largo1 >= Li_Largo2
                    DoEvents
                    
                    'Llena_Datos
                    ls_convenio = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_convenio) + 1
                    
'************************' validacion de campo con_casilla ****************
'Modificado por        : Entersoft
'Fecha de modificacion : 13/08/2009
'Detalle               : Validacion de campo con_casilla en tabla convenio
                    
                    If Trim(ls_convenio) = "0" Then
                        ls_convenio = "00"
                        Else
                        If Trim(ls_convenio) = "1" Then
                            ls_convenio = "01"
                            Else
                            If Trim(ls_convenio) = "2" Then
                                ls_convenio = "02"
                                Else
                                ls_convenio = "00"
                            End If
                         End If
                    End If
'***************************************************************************

                    ls_rendicion = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_rendicion) + 1
                    
'*********************** validacion de campo con_entrega_vv ****************
'Modificado por        : Entersoft
'Fecha de modificacion : 21/08/2009
'Detalle               : Validacion de campo con_entrega_vv en tabla convenio
                    
                    If Trim(ls_rendicion) <> "S" And Trim(ls_rendicion) <> "N" Then
                        ls_rendicion = "S"
                    End If
'***************************************************************************
                    ls_contin = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_contin) + 1

'*********************** validacion de campo anx_con_contingencia ****************
'Modificado por        : Entersoft
'Fecha de modificacion : 05/10/2009
'Detalle               : Validacion de campo anx_con_contingencia en tabla anx_convenio
                    
                    If ls_contin <> "S" And ls_contin <> "N" Then
                       ls_contin = "N"
                    End If
'***************************************************************************
                    ls_ctacte = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte) + 1
                    ls_glosa = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_glosa) + 1

'***************************************************************************
'Observaciones : valida campos nuevos
'Desarrollo    : Entersoft
'Fecha         : 10-09-2009

                    ls_nom_convenio = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_nom_convenio) + 1
                    
                    ls_vigencia = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_vigencia) + 1
                    
                    ls_fecha_in = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fecha_in) + 1
                
                    ls_fecha_fin = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fecha_fin) + 1

                    ls_ofi_rela = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ofi_rela) + 1

                    ls_ofi_rendi = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ofi_rendi) + 1

                    ls_ofi_contrato = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ofi_contrato) + 1

                    ls_ctacte_interna = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte_interna) + 1

                    ls_ctacte_empresa = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte_empresa) + 1

                    ls_ctacte_factu = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte_factu) + 1

                    ls_act_economica = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_act_economica) + 1
                    
                    If ls_act_economica = "A1" Then
                       ls_act_economica = "Fisco y Rep. Gubernativas"
                       Else
                       If ls_act_economica = "A2" Then
                          ls_act_economica = "Agencias Descentralizadas"
                          Else
                          If ls_act_economica = "A3" Then
                              ls_act_economica = "Sistema Previsional"
                              Else
                              If ls_act_economica = "A4" Then
                              ls_act_economica = "Empresas Publicas"
                                  Else
                                  If ls_act_economica = "A5" Then
                                  ls_act_economica = "Gran Mineria Cobre y Andi"
                                      Else
                                      If ls_act_economica = "A6" Then
                                      ls_act_economica = "Municipalidades"
                                          Else
                                          If ls_act_economica = "B1" Then
                                          ls_act_economica = "Personas Juridicas"
                                              Else
                                              If ls_act_economica = "BC" Then
                                              ls_act_economica = "Persona Natural"
                                                  Else
                                                  If ls_act_economica = "BL" Then
                                                  ls_act_economica = "Sistema Previsional"
                                                      Else
                                                      If ls_act_economica = "C1" Then
                                                      ls_act_economica = "Banco Central de Chile"
                                                          Else
                                                          If ls_act_economica = "C2" Then
                                                          ls_act_economica = "Banco del Estado de Chile"
                                                              Else
                                                              If ls_act_economica = "C3" Then
                                                              ls_act_economica = "Otros"
                                                                  Else
                                                                  If ls_act_economica = "D1" Then
                                                                  ls_act_economica = "Bancos Comerciales"
                                                                      Else
                                                                      If ls_act_economica = "D2" Then
                                                                      ls_act_economica = "Bancos Fomento e Ificoop"
                                                                          Else
                                                                          If ls_act_economica = "D3" Then
                                                                          ls_act_economica = "Soc. Financ. Ahorro y Cred."
                                                                              Else
                                                                              If ls_act_economica = "D4" Then
                                                                              ls_act_economica = "Sis. Nac. Ahorro y Presta."
                                                                                  Else
                                                                                  If ls_act_economica = "D5" Then
                                                                                  ls_act_economica = "Bancos del Exterior"
                                                                                      Else
                                                                                      If ls_act_economica = "D6" Then
                                                                                      ls_act_economica = "Otras Entid. Financieras"
                                                                                      End If
                                                                                   End If
                                                                              End If
                                                                          End If
                                                                      End If
                                                                  End If
                                                              End If
                                                          End If
                                                      End If
                                                  End If
                                              End If
                                          End If
                                      End If
                                  End If
                              End If
                          End If
                       End If
                    End If
        
                    ls_adm_rechazo = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_adm_rechazo) + 1
                    
                    If ls_adm_rechazo = "00" Then
                       ls_adm_rechazo = "Sin Administracion Rechazos"
                       Else
                       If ls_adm_rechazo = "01" Then
                       ls_adm_rechazo = "VV Beneficiario"
                           Else
                           If ls_adm_rechazo = "02" Then
                           ls_adm_rechazo = "VV Empresa por monto rechazado"
                               Else
                               If ls_adm_rechazo = "03" Then
                               ls_adm_rechazo = "Abono a Cta.Cte. Empresa"
                                   Else
                                   If ls_adm_rechazo = "04" Then
                                   ls_adm_rechazo = "VV Beneficiario"
                                   End If
                               End If
                           End If
                       End If
                    End If
                
                    ls_contrata_avb = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_contrata_avb) + 1

                    If ls_contrata_avb = "S" Then
                        ls_contrata_avb = "SI"
                        Else
                        If ls_contrata_avb = "N" Then
                           ls_contrata_avb = "NO"
                        End If
                    End If
                    
'                   ls_concilia = trae_campo(ls_ret, Li_Largo2)  Campo Rendicion contratada
                    ls_concilia = recupera_canal()
                    Li_Largo2 = Li_Largo2 + Len(ls_concilia) + -1
                    If ls_concilia = "S" Then
                       ls_concilia = "SI"
                       Else
                       If ls_concilia = "N" Then
                          ls_concilia = "NO"
                       End If
                    End If
                    
                    ls_nom_contacto = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_nom_contacto) + 1

                    ls_fono_contacto = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fono_contacto) + 1

                    ls_fax_contacto = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fax_contacto) + 1

                    ls_visado = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_visado) + 1

                    If ls_visado = "M" Then
                        ls_visado = "Modificacion por visar"
                        Else
                        If ls_visado = "N" Then
                           ls_visado = "Visado OK"
                           Else
                           If ls_visado = "C" Then
                              ls_visado = "Creacion de Visado"
                              Else
                              If ls_visado = "P" Then
                                 ls_visado = "Pendiente de envio a Visar"
                                 Else
                                 If ls_visado = "T" Then
                                    ls_visado = "Cierre de Visado"
                                    Else
                                    If ls_visado = "I" Then
                                       ls_visado = "Creacion Incompleta"
                                    End If
                                 End If
                              End If
                           End If
                        End If
                    End If
'***************************************************************************
                    
                    If ls_glosa = "convenio" Then
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
    
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Tipo de Convenio  "
                        
                        If (ls_convenio <> "") And (ls_convenio = "00") Then ls_convenio = SacaValorINI("TipoConvenio", "00", DIRECTORIO_01 + ARCHIVO_INI01)
                        If (ls_convenio <> "") And (ls_convenio = "01") Then ls_convenio = SacaValorINI("TipoConvenio", "01", DIRECTORIO_01 + ARCHIVO_INI01)
                        If (ls_convenio <> "") And (ls_convenio = "02") Then ls_convenio = SacaValorINI("TipoConvenio", "02", DIRECTORIO_01 + ARCHIVO_INI01)
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = ls_convenio
                       
                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Tipo de Rendición  "
                        
'*********************** validacion de campo con_entrega_vv ****************
'Modificado por        : Entersoft
'Fecha de modificacion : 21/08/2009
'Detalle               : Validacion de campo con_entrega_vv en tabla convenio
                        
                    '  If (ls_rendicion <> "") And (ls_rendicion = "S") Then ls_rendicion = SacaValorINI("TipoRendicion", "S", DIRECTORIO_01 + ARCHIVO_INI01)
                    '  If (ls_rendicion <> "") And (ls_rendicion = "N") Then ls_rendicion = SacaValorINI("TipoRendicion", "N", DIRECTORIO_01 + ARCHIVO_INI01)
                        
                        If (ls_rendicion = "S") Then ls_rendicion = SacaValorINI("TipoRendicion", "S", DIRECTORIO_01 + ARCHIVO_INI01)
                        If (ls_rendicion = "N") Then ls_rendicion = SacaValorINI("TipoRendicion", "N", DIRECTORIO_01 + ARCHIVO_INI01)

'***************************************************************************
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = ls_rendicion
                        
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        'Siguiente linea.
                        '****************
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Contingencia  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_contin)
    
                        
                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta. Cte. Rendición  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte), "0-000-00000-00")
                        
'***************************************************************************
'Observaciones : Agrega linea incorporando campos nuevos
'Desarrollo    : Entersoft
'Fecha         : 10-09-2009
                        'Siguiente linea. //ras
                        '****************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Nombre Convenio  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_nom_convenio)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Convenio Vigente  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_vigencia)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
         
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fecha Creacion  "
                                      
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format((ls_fecha_in), "dd/mm/yyyy")

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fecha Vencimiento  "
    
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format((ls_fecha_fin), "dd/mm/yyyy")
        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Of.Relacionadora  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_ofi_rela)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Of.Rendidora  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_ofi_rendi)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Oficina Contrato  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_ofi_contrato)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta.Cte.Interna  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte_interna), "0-000-00000-00")
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta.Cte.Pago  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte_empresa), "0-000-00000-00")
                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta.Cte.Facturación  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte_factu), "0-000-00000-00")
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Personería Jurídica  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_act_economica)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Adm.Rechazos  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_adm_rechazo)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Contrata Aviso  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_contrata_avb)

                         GrillaVisado.Col = 3
                         GrillaVisado.Row = li_filo
                         GrillaVisado.Text = "Conciliación  "
      
                         GrillaVisado.Col = 4
                         GrillaVisado.Row = li_filo
                         GrillaVisado.Text = Trim(ls_concilia)
          
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Contacto Operativo  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_nom_contacto)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fono Contacto  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_fono_contacto)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fax Contacto  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_fax_contacto)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Estado Visado  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_visado)
                        
                        
'***************************************************************************
                    End If
                Loop
                
            End If
        Loop
        
      'End If
    
    FiltroTabDesp.Close
    Call ColorearGrillaProductos
    
  Else
    Retorno% = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
    Retorno% = PoneValorINI("Parametros", "LlaveVisad", Data, DIRECTORIO_01 + ARCHIVO_INI01)
    Call EjecutarPrograma(NombreArchivo, 1)
    If SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01) = "OK" Then
      Fila = 1
      NombreArchivo = DIRECTORIO_01 + Mid$(ARCHIVO_TXT02, 1, 3) + Producto + Mid$(ARCHIVO_TXT02, 7)
      Open NombreArchivo For Input As #1
      While Not EOF(1)
        Input #1, Linea
        Columna = 1
        ' FI2026 - .Rows - 1 = filas de datos (era .MaxRows)
        If Fila > GrillaVisado.Rows - 1 Then
          GrillaVisado.Rows = Fila + 1 ' FI2026 - era .MaxRows = Fila
        End If
        While InStr(Linea, "¬") <> 0
          GrillaVisado.Col = Columna
          GrillaVisado.Row = Fila
          GrillaVisado.Text = Mid$(Linea, 1, InStr(Linea, "¬") - 1)
          Linea = Mid$(Linea, InStr(Linea, "¬") + 1)
          
          If InStr(Linea, ";") Then
              Mid$(Linea, InStr(Linea, ";"), 1) = ","
          End If
          Columna = Columna + 1
        Wend
        GrillaVisado.Col = Columna
        GrillaVisado.Row = Fila
        GrillaVisado.Text = Linea
        Fila = Fila + 1
      Wend
      'If Producto = "SP0" Then
        GrillaVisado.Col = 2
        GrillaVisado.Row = 1
        gCodConvenio = GrillaVisado.Text
        
        GrillaVisado.Rows = GrillaVisado.Rows + 1
        li_filo = GrillaVisado.Rows - 1 ' FI2026 - era .MaxRows: cantidad de filas de DATOS

        inicio = 1
        sigue = True
        ''Query de Consulta por Medio de Pago y Operaciones
        ''*************************************************
        'RutCl = DesformateaRut(Visando.RutCliente.Caption)
        GRut_SForma = DesformateaRut(Visando.RutCliente.Caption)
        
        ls_sql = ""
'*************************************************************************************************************************************************************************
'Modificado por        : Entersoft
'Fecha de modificacion : 13/08/2009
'Detalle               : Reemplazar en aplicacion Visado campo anx_con_tipo_conv de tabla anx_convenio por campo con_casilla de tabla convenio
'Fecha de modificacion : 21/08/2009
'                        Reemplazar en aplicacion visado campo anx_con_tipo_ren de tabla anx_convenio por campo con_entrega_vv de tabla convenio

'        ls_sql = ls_sql & "select anx_con_tipo_conv + '~' + anx_con_tipo_rend + '~' + anx_con_contingencia + '~' + isnull(anx_con_cta_rend, '0')  + '~' + 'convenio'+ '~'  "
'        ls_sql = ls_sql & "select con_casilla + '~' + con_entrega_vv + '~' + anx_con_contingencia + '~' + isnull(anx_con_cta_rend, '0')  + '~' + 'convenio'+ '~' + con_nom_convenio + '~'"
         ls_sql = ls_sql & "select con_casilla + '~' + con_entrega_vv  + '~' + anx_con_contingencia + '~' + isnull(anx_con_cta_rend, '0')  + '~' + 'convenio'+ '~' +"
         ls_sql = ls_sql & "con_nom_convenio + '~'   + con_vigencia        + '~' + "
         ls_sql = ls_sql & "convert(char(10),con_fecha_in,103)  + '~' + convert(char(10),con_fecha_fin,103) + '~' + "
         ls_sql = ls_sql & "con_ofi_rela  + '~' + con_ofi_rendi + '~' + "    'campos nuevos
         ls_sql = ls_sql & "con_ofi_contrato + '~'  + con_ctacte_interna + '~' + con_ctacte_empresa + '~' + con_ctacte_factu  + '~' + con_act_economica + '~' + con_adm_rechazo + '~' + "                          'campos nuevos
         ls_sql = ls_sql & "con_contrata_avb + '~'  + con_nom_contacto    + '~' + con_fono_contacto + '~' + con_fax_contacto  + '~' + con_visado + '~'  "                                                          'campos nuevos
'       ls_sql = ls_sql & "con_contrata_avb + '~'  + rendicion          + '~' + con_nom_contacto    + '~' + con_fono_contacto + '~' + con_fax_contacto  + '~' + con_visado + '~'
'       ls_sql = ls_sql & " from anx_convenio "
        ls_sql = ls_sql & " from convenio, anx_convenio "
'       ls_sql = ls_sql & " where anx_con_rut_empresa = " & Mid(GRut_SForma, 1, Len(GRut_SForma) - 1) & " "
        ls_sql = ls_sql & " where con_cliente = " & Mid(GRut_SForma, 1, Len(GRut_SForma) - 1) & " "
        ls_sql = ls_sql & " and con_convenio = '" & gCodConvenio & "'"
        ls_sql = ls_sql & " and con_cliente *= anx_con_rut_empresa "
        ls_sql = ls_sql & " and con_convenio *= anx_con_convenio "
      
'*************************************************************************************************************************************************************************



        Do While sigue
            CadenaDatos = ""
            DoEvents
            CadenaDatos = Llamada_ConReintentar("U076", "GS23", "b12256781234R00000011npagos      ", ls_sql, "8K")
            ls_existe = Mid(CadenaDatos, 7, 1)
            
            If ls_existe = "N" Then sigue = False
            'ls_ret = Mid(CadenaDatos, 15, Len(CadenaDatos))
            ls_ret = Mid(CadenaDatos, 19, Len(CadenaDatos))
            If ls_ret <> "" Then
                Li_Largo1 = Len(ls_ret)
                Li_Largo2 = 1
                
                Do While Li_Largo1 >= Li_Largo2
                    DoEvents
                    
                    'Llena_Datos
                    ls_convenio = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_convenio) + 1

'************************' validacion de campo con_casilla ***************
'Modificado por        : Entersoft
'Fecha de modificacion : 13/08/2009
'Detalle               : Validacion de campo con_casilla en tabla convenio
                    
                    If Trim(ls_convenio) = "0" Then
                        ls_convenio = "00"
                        Else
                        If Trim(ls_convenio) = "1" Then
                            ls_convenio = "01"
                            Else
                            If Trim(ls_convenio) = "2" Then
                                ls_convenio = "02"
                                Else
                                ls_convenio = "00"
                            End If
                         End If
                    End If
'*************************************************************************
                    ls_rendicion = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_rendicion) + 1

'*********************** validacion de campo con_entrega_vv ****************
'Modificado por        : Entersoft
'Fecha de modificacion : 21/08/2009
'Detalle               : Validacion de campo con_entrega_vv en tabla convenio
                    
                    If Trim(ls_rendicion) <> "S" And Trim(ls_rendicion) <> "N" Then
                        ls_rendicion = "S"
                    End If
'***************************************************************************
                 
                    ls_contin = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_contin) + 1
'*********************** validacion de campo anx_con_contingencia ****************
'Modificado por        : Entersoft
'Fecha de modificacion : 05/10/2009
'Detalle               : Validacion de campo anx_con_contingencia en tabla anx_convenio
                    
                    If ls_contin <> "S" And ls_contin <> "N" Then
                       ls_contin = "N"
                    End If
'***************************************************************************
                    
                    ls_ctacte = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte) + 1
                    ls_glosa = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_glosa) + 1

'***************************************************************************
'Observaciones : valida campos nuevos
'Desarrollo    : Entersoft
'Fecha         : 10-09-2009

                    ls_nom_convenio = trae_campo(ls_ret, Li_Largo2)  'nombre convenio
                    Li_Largo2 = Li_Largo2 + Len(ls_nom_convenio) + 1
                    
                    ls_vigencia = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_vigencia) + 1
                    
                    If ls_vigencia = "N" Then
                       ls_vigencia = "NO"
                       Else
                        If ls_vigencia = "S" Then
                           ls_vigencia = "SI"
                        End If
                    End If
                    
                    ls_fecha_in = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fecha_in) + 1
                    
                    ls_fecha_fin = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fecha_fin) + 1
                    
                    ls_ofi_rela = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ofi_rela) + 1

                    ls_ofi_rendi = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ofi_rendi) + 1

                    ls_ofi_contrato = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ofi_contrato) + 1

                    ls_ctacte_interna = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte_interna) + 1

                    ls_ctacte_empresa = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte_empresa) + 1

                    ls_ctacte_factu = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_ctacte_factu) + 1

                    ls_act_economica = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_act_economica) + 1
                    
                    If ls_act_economica = "A1" Then
                       ls_act_economica = "Fisco y Rep. Gubernativas"
                       Else
                       If ls_act_economica = "A2" Then
                          ls_act_economica = "Agencias Descentralizadas"
                          Else
                          If ls_act_economica = "A3" Then
                              ls_act_economica = "Sistema Previsional"
                              Else
                              If ls_act_economica = "A4" Then
                              ls_act_economica = "Empresas Publicas"
                                  Else
                                  If ls_act_economica = "A5" Then
                                  ls_act_economica = "Gran Mineria Cobre y Andi"
                                      Else
                                      If ls_act_economica = "A6" Then
                                      ls_act_economica = "Municipalidades"
                                          Else
                                          If ls_act_economica = "B1" Then
                                          ls_act_economica = "Personas Juridicas"
                                              Else
                                              If ls_act_economica = "BC" Then
                                              ls_act_economica = "Persona Natural"
                                                  Else
                                                  If ls_act_economica = "BL" Then
                                                  ls_act_economica = "Sistema Previsional"
                                                      Else
                                                      If ls_act_economica = "C1" Then
                                                      ls_act_economica = "Banco Central de Chile"
                                                          Else
                                                          If ls_act_economica = "C2" Then
                                                          ls_act_economica = "Banco del Estado de Chile"
                                                              Else
                                                              If ls_act_economica = "C3" Then
                                                              ls_act_economica = "Otros"
                                                                  Else
                                                                  If ls_act_economica = "D1" Then
                                                                  ls_act_economica = "Bancos Comerciales"
                                                                      Else
                                                                      If ls_act_economica = "D2" Then
                                                                      ls_act_economica = "Bancos Fomento e Ificoop"
                                                                          Else
                                                                          If ls_act_economica = "D3" Then
                                                                          ls_act_economica = "Soc. Financ. Ahorro y Cred."
                                                                              Else
                                                                              If ls_act_economica = "D4" Then
                                                                              ls_act_economica = "Sis. Nac. Ahorro y Presta."
                                                                                  Else
                                                                                  If ls_act_economica = "D5" Then
                                                                                  ls_act_economica = "Bancos del Exterior"
                                                                                      Else
                                                                                      If ls_act_economica = "D6" Then
                                                                                      ls_act_economica = "Otras Entid. Financieras"
                                                                                      End If
                                                                                   End If
                                                                              End If
                                                                          End If
                                                                      End If
                                                                  End If
                                                              End If
                                                          End If
                                                      End If
                                                  End If
                                              End If
                                          End If
                                      End If
                                  End If
                              End If
                          End If
                       End If
                    End If
        
                    ls_adm_rechazo = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_adm_rechazo) + 1
                    
                    If ls_adm_rechazo = "00" Then
                       ls_adm_rechazo = "Sin Administracion Rechazos"
                       Else
                       If ls_adm_rechazo = "01" Then
                       ls_adm_rechazo = "VV Beneficiario"
                           Else
                           If ls_adm_rechazo = "02" Then
                           ls_adm_rechazo = "VV Empresa por monto rechazado"
                               Else
                               If ls_adm_rechazo = "03" Then
                               ls_adm_rechazo = "Abono a Cta.Cte. Empresa"
                                   Else
                                   If ls_adm_rechazo = "04" Then
                                   ls_adm_rechazo = "VV Beneficiario"
                                   End If
                               End If
                           End If
                       End If
                    End If
                
                    ls_contrata_avb = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_contrata_avb) + 1

                    If ls_contrata_avb = "S" Then
                        ls_contrata_avb = "SI"
                        Else
                        If ls_contrata_avb = "N" Then
                           ls_contrata_avb = "NO"
                        End If
                    End If

'                   ls_concilia = trae_campo(ls_ret, Li_Largo2)  Campo Rendicion contratada
                    ls_concilia = recupera_canal()
                    Li_Largo2 = Li_Largo2 + Len(ls_concilia) + -1
                    If ls_concilia = "S" Then
                       ls_concilia = "SI"
                       Else
                       If ls_concilia = "N" Then
                          ls_concilia = "NO"
                       End If
                    End If

                    ls_nom_contacto = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_nom_contacto) + 1

                    ls_fono_contacto = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fono_contacto) + 1

                    ls_fax_contacto = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_fax_contacto) + 1

                    ls_visado = trae_campo(ls_ret, Li_Largo2)
                    Li_Largo2 = Li_Largo2 + Len(ls_visado) + 1

                    If ls_visado = "M" Then
                        ls_visado = "Modificacion por visar"
                        Else
                        If ls_visado = "N" Then
                           ls_visado = "Visado OK"
                           Else
                           If ls_visado = "C" Then
                              ls_visado = "Creacion de Visado"
                              Else
                              If ls_visado = "P" Then
                                 ls_visado = "Pendiente de envio a Visar"
                                 Else
                                 If ls_visado = "T" Then
                                    ls_visado = "Cierre de Visado"
                                    Else
                                    If ls_visado = "I" Then
                                       ls_visado = "Creacion Incompleta"
                                    End If
                                 End If
                              End If
                           End If
                        End If
                    End If
'***************************************************************************
                    If ls_glosa = "convenio" Then
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
    
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Tipo de Convenio  "
                        
                        If (ls_convenio <> "") And (ls_convenio = "00") Then ls_convenio = SacaValorINI("TipoConvenio", "00", DIRECTORIO_01 + ARCHIVO_INI01)
                        If (ls_convenio <> "") And (ls_convenio = "01") Then ls_convenio = SacaValorINI("TipoConvenio", "01", DIRECTORIO_01 + ARCHIVO_INI01)
                        If (ls_convenio <> "") And (ls_convenio = "02") Then ls_convenio = SacaValorINI("TipoConvenio", "02", DIRECTORIO_01 + ARCHIVO_INI01)
                        
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = ls_convenio
    
                        
                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Tipo de Rendición  "
                        
'*********************** validacion de campo con_entrega_vv ****************
'Modificado por        : Entersoft
'Fecha de modificacion : 21/08/2009
'Detalle               : Validacion de campo con_entrega_vv en tabla convenio
                        
                       'If (ls_rendicion <> "") And (ls_rendicion = "S") Then ls_rendicion = SacaValorINI("TipoRendicion", "S", DIRECTORIO_01 + ARCHIVO_INI01)
                       'If (ls_rendicion <> "") And (ls_rendicion = "N") Then ls_rendicion = SacaValorINI("TipoRendicion", "N", DIRECTORIO_01 + ARCHIVO_INI01)
                        
                        
                        If (ls_rendicion = "S") Then ls_rendicion = SacaValorINI("TipoRendicion", "S", DIRECTORIO_01 + ARCHIVO_INI01)
                        If (ls_rendicion = "N") Then ls_rendicion = SacaValorINI("TipoRendicion", "N", DIRECTORIO_01 + ARCHIVO_INI01)

'***************************************************************************
                        
                        
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = ls_rendicion
                        
                        
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        'Siguiente linea.
                        '****************
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Contingencia  "
                        
                        If ls_contin = "S" Then ls_contin = "SI"
                        If ls_contin = "N" Then ls_contin = "NO"
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_contin)
    
                        
                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta. Cte. Rendición  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte), "0-000-00000-00")
                        
'***************************************************************************
'Observaciones : Agrega linea incorporando campos nuevos
'Desarrollo    : Entersoft
'Fecha         : 10-09-2009
                        'Siguiente linea. //ras
                        '****************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Nombre Convenio  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_nom_convenio)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Convenio Vigente  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_vigencia)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                         
                        GrillaVisado.Col = 1
                  
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fecha Creacion  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format((ls_fecha_in), "dd/mm/yyyy")
 
                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fecha Vencimiento  "
     
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format((ls_fecha_fin), "dd/mm/yyyy")
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Of.Relacionadora  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_ofi_rela)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Of.Rendidora  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_ofi_rendi)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Oficina Contrato  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_ofi_contrato)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta.Cte.Interna  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte_interna), "0-000-00000-00")
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta.Cte.Pago  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte_empresa), "0-000-00000-00")
                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Cta.Cte.Facturación  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Format(Trim(ls_ctacte_factu), "0-000-00000-00")
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Personería Jurídica  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_act_economica)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Adm.Rechazos  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_adm_rechazo)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Contrata Aviso  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_contrata_avb)

                         GrillaVisado.Col = 3
                         GrillaVisado.Row = li_filo
                         GrillaVisado.Text = "Conciliación  "
      
                         GrillaVisado.Col = 4
                         GrillaVisado.Row = li_filo
                         GrillaVisado.Text = Trim(ls_concilia)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Contacto Operativo  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_nom_contacto)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fono Contacto  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_fono_contacto)
                        
                        '***************
                        li_filo = li_filo + 1
                        GrillaVisado.Rows = li_filo + 1
                        
                        GrillaVisado.Col = 1
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Fax Contacto  "
                        
                        GrillaVisado.Col = 2
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_fax_contacto)

                        GrillaVisado.Col = 3
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = "Estado Visado  "
                        
                        GrillaVisado.Col = 4
                        GrillaVisado.Row = li_filo
                        GrillaVisado.Text = Trim(ls_visado)
                        
'***************************************************************************
                    End If
                Loop
                
            End If
        Loop
        
      'End If
      Close #1
      Kill NombreArchivo
      Call ColorearGrillaProductos
      


    Else
      Unload Me
    End If
  End If
End Sub

Function trae_campo(Data_Resp As String, Lgi_Largo As Integer) As String
    Dim Lgc_Campo As String
    Dim i         As Integer

    Lgc_Campo$ = ""
    For i = Lgi_Largo To Len(Data_Resp$)
        If Mid$(Data_Resp$, i, 1) = "~" Then
           Exit For
        End If
       
        Lgc_Campo$ = Lgc_Campo$ & Mid$(Data_Resp$, i, 1)
    Next i

    trae_campo = Lgc_Campo$
End Function

Private Sub Aceptar_Click()
  Call MostrarPautaVisado(False)
  If HabilitarAprobacionVisado() Then
    PantallaVisar.Aprobar.Enabled = True
  End If
End Sub
Private Sub Aceptar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Aceptar_Click
  End If
End Sub

Private Sub Aprobar_Click()
  Call ColocarTickGrillaProductos(Aprobado, COLOR_03)
  Call HabilitarBotonTransmision
  Call Cancelar_Click
End Sub
Private Sub Aprobar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub BtnPauta_Click(Index As Integer)
Dim Indice As Integer
  Screen.MousePointer = 11
  If EjecutarActualBoton(Trim$(Mid(BtnPauta(Index).Tag, 3))) Then
    Indice = Val(Mid(BtnPauta(Index).Tag, 2, 1))
    If LlamadaProgramaAsociado(Indice, 1) Then
      Glosa(Indice).Enabled = False
      BtnPauta(Index).Enabled = False
    End If
  End If
  Screen.MousePointer = 0
End Sub
Private Sub BtnPauta_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Aceptar_Click
  End If
End Sub
Private Sub Cancelar_Click()
  Unload Me
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub
Private Sub ChkPauta_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub

Private Sub Form_Activate()
  Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
  Screen.MousePointer = 11
  With Visando
    .Panel3D3.Enabled = False
    .Panel3D4.Enabled = False
    Rut.Caption = .RutCliente.Caption
    Nombre.Caption = .NombreCliente.Caption
    .GrillaProductos.Row = .GrillaProductos.Row
    .GrillaProductos.Col = 3
    Producto.Caption = .GrillaProductos.Text
  End With
  Call CentrarVentana(Me)
  With GrillaVisado
    ' FI2026 - MaxCols=4 (Spread: columnas de DATOS 1..4) -> Cols=5
    .Cols = 5
    ' FI2026 - se oculta la fila fija 0 de MSFlexGrid (el Spread original no
    ' mostraba encabezado de columna). La COLUMNA 0 queda visible a proposito:
    ' es donde ColocarTickGrillaVisado pone el icono (.Col = 0 + .CellPicture).
    .RowHeight(0) = 0
    .ColWidth(1) = 1920  ' FI2026 - eran 16 caracteres (Spread)
    .ColWidth(2) = 2160  ' FI2026 - eran 18 caracteres (Spread)
    .ColWidth(3) = 1800  ' FI2026 - eran 15 caracteres (Spread)
    .ColWidth(4) = 2640  ' FI2026 - eran 22 caracteres (Spread)
  End With
  Call MuestraDetalleEnGrilla
  Visando.Enabled = False
  If Not Recorte Then
    If CrearPautaValidacion() Then
      Call HabilitarPautaVisado(True)
    Else
      Call HabilitarPautaVisado(False)
    End If
  Else
    Unload Me
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  With Visando
    .Panel3D3.Enabled = True
    .Panel3D4.Enabled = True
    .Enabled = True
    If .GrillaProductos.BackColor <> COLOR_04 Then
      .GrillaProductos.Col = 1
      ' FI2026 SHG - OperationMode (Spread) sin equivalente en MSFlexGrid
      '.GrillaProductos.OperationMode = 0
    End If
    .GrillaProductos.SetFocus
  End With
End Sub
Private Sub ColorearGrillaProductos()
  For i% = 1 To GrillaVisado.Rows - 1 ' FI2026 - era .MaxRows
    GrillaVisado.Row = i%
    For j% = 1 To 4
      GrillaVisado.Col = j%
      GrillaVisado.BackColor = COLOR_05
      If j% Mod 2 = 0 Then
        GrillaVisado.ForeColor = COLOR_01
      End If
    Next j%
  Next i%
End Sub
Private Function CrearPautaValidacion() As Boolean
Dim i As Integer
Dim Tope As Integer
Dim FiltroTabPta As Recordset
Dim IndiceMayorBoton As Integer
Dim IndiceMayorCheck As Integer
Dim SeUsoPrimerBoton As Boolean
Dim SeUsoPrimerCheck As Boolean
  CrearPautaValidacion = False
  Visando.GrillaProductos.Col = 4
  Set FiltroTabPta = BaseVisado.OpenRecordset("SELECT * FROM Tab_Pta WHERE Pta_Prod='" + _
                                                                                 Visando.GrillaProductos.Text + "' ORDER BY Pta_Orden ASC;", dbOpenSnapshot)
  If FiltroTabPta.RecordCount > 0 Then
    i = 0
    Tope = 110
    IndiceMayorBoton = 0
    IndiceMayorCheck = 0
    SeUsoPrimerBoton = False
    SeUsoPrimerCheck = False
    CrearPautaValidacion = True
    IndiceBotonParaVerFirma = -1
    While Not FiltroTabPta.EOF
      With Glosa(i)
        .Left = 720
        .Top = Tope
        .Caption = FiltroTabPta("Pta_Glosa")
        .Visible = True
        'Panel.Height = .Top + .Height + 30
      End With
      With Numero(i)
        .Left = 90
        .Top = Tope
        .Caption = LTrim$(Str(i + 1)) + "."
        .Visible = True
      End With
      Select Case FiltroTabPta("Pta_Check")
        Case 1
          If SeUsoPrimerCheck Then
            IndiceMayorCheck = IndiceMayorCheck + 1
            Load Pauta
          End If
          With Pauta
            .Left = 450
            .Top = Tope
            .Visible = True
          End With
          SeUsoPrimerCheck = True
        Case 2
          If SeUsoPrimerBoton Then
            IndiceMayorBoton = IndiceMayorBoton + 1
            Load BtnPauta(IndiceMayorBoton)
          End If
          With BtnPauta(IndiceMayorBoton)
            .Tag = Format$(FiltroTabPta("Pta_Prior"), "0") + LTrim$(Str(FiltroTabPta("Pta_Orden")))
            If Not IsNull(FiltroTabPta("Pta_Depende")) Then
              If Trim$(FiltroTabPta("Pta_Depende")) <> "" Then
                .Tag = .Tag + FiltroTabPta("Pta_Depende")
              End If
            End If
            .Left = 450
            .Top = Tope
            .Visible = True
          End With
          SeUsoPrimerBoton = True
          If FiltroTabPta("Pta_VerFirma") Then
            IndiceBotonParaVerFirma = IndiceMayorBoton
          End If
      End Select
      FiltroTabPta.MoveNext
      Tope = Glosa(i).Top + Glosa(i).Height + 30
      i = i + 1
      Load Glosa(i)
      Load Numero(i)
    Wend
  End If
  FiltroTabPta.Close
End Function
Private Sub GrillaVisado_GotFocus()
  GrillaVisado.BorderStyle = 1
End Sub

Private Sub GrillaVisado_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub
Private Sub GrillaVisado_LostFocus()
  GrillaVisado.BorderStyle = 0
End Sub

Private Sub Imprimir_Click()
Dim Cadena As String
Dim LargoGlosaMasLarga As Integer

  Screen.MousePointer = 11
On Error GoTo ErrorDeImpresion
  Printer.FontName = "Arial"
  Printer.FontName = "Courier New"
  Printer.FontSize = "10"
  Printer.Print "Banco de Chile"; Tab(60); "Fecha Emisión : " + Format$(Date, "dd/mmm/yyyy")
  Printer.Print Tab(60); "Hora Emisión  : " + Format$(Time, "hh:mm:ss")
  Printer.Print
  Printer.Print
  Printer.Print Spc(26); "INFORME DEL PRODUCTO A VISAR"
  Printer.Print Spc(26); String(28, "-")
  Printer.Print
  Printer.Print Tab(10); "Rut Cliente"; Tab(27); ": " + Rut.Caption
  Printer.Print Tab(10); "Nombre Cliente"; Tab(27); ": " + Nombre.Caption
  Printer.Print Tab(10); "Producto"; Tab(27); ": " + Producto.Caption
  Printer.Print Tab(10); "Acción"; Tab(27); ": " + Accion.Caption
  Printer.Print Tab(10); "Oficina"; Tab(27); ": " + Oficina
  With Visando
    .GrillaProductos.Col = 1
    Printer.Print Tab(10); "Ejecutivo"; Tab(27); ": " + .GrillaProductos.Text
  End With
  Printer.Print
  LargoGlosaMasLarga = 0
  With GrillaVisado
    For i% = 1 To .Rows - 1 ' FI2026 - era .MaxRows
      .Row = i%
      For j% = 1 To .Cols - 1 ' FI2026 - era .MaxCols
        If j% Mod 2 <> 0 Then
          .Col = j%
          If Len(Trim(.Text)) > LargoGlosaMasLarga Then
            LargoGlosaMasLarga = Len(Trim(.Text))
          End If
        End If
      Next j%
    Next i%
    For i% = 1 To .Rows - 1 ' FI2026 - era .MaxRows
      .Row = i%
      For j% = 1 To .Cols - 1 ' FI2026 - era .MaxCols
        .Col = j%
        If j% Mod 2 <> 0 Then
          Cadena = Space(2) + Trim$(.Text) + String(LargoGlosaMasLarga - Len(Trim(.Text)), " ")
        Else
          If (Trim$(Cadena) <> "") Or (Trim$(.Text) <> "") Then
            If Trim$(Cadena) <> "" Then
              Printer.Print Tab(10); Cadena + " : " + Trim$(.Text)
            Else
              Printer.Print Tab(10); Cadena + Space(3) + Trim$(.Text)
            End If
          End If
        End If
      Next j%
    Next i%
  End With
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
  If Pauta.Enabled Then
    Pauta.SetFocus
  Else
    If Aprobar.Enabled Then
      Aprobar.SetFocus
    Else
      Cancelar.SetFocus
    End If
  End If
  Screen.MousePointer = 0
End Sub
Private Sub Imprimir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub

Private Sub Pauta_Click()
  Call MostrarPautaVisado(True)
End Sub
Private Sub MostrarPautaVisado(ByVal Estado As Boolean)
  Pauta.Enabled = Not (Estado)
  Cancelar.Enabled = Not (Estado)
  If (Estado) And (Aprobar.Enabled) Then
    Aprobar.Enabled = False
  End If
  Panel.Visible = Estado
  Aceptar.Visible = Estado
  '*Sombra(1).Visible = Estado
  '*If Estado Then
  '*  Sombra(1).ZOrder 1
  '*End If
  'Sombra(0).Visible = Estado
  'VisorPauta.Visible = Estado
  Panel.Visible = Estado
  GrillaVisado.Enabled = Not (Estado)
End Sub
Private Sub Pauta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call Cancelar_Click
  End If
End Sub

Private Sub HabilitarPautaVisado(ByVal Estado As Boolean)
  With PantallaVisar
    .Pauta.Enabled = Estado
    .Aprobar.Enabled = Not (Estado)
  End With
End Sub
Private Function LlamadaProgramaAsociado(ByVal CorrelativoSegunLinea As String, ByVal Tipo As String) As Boolean
Dim Retorno As Integer
Dim Producto As String
Dim Parametro As String
Dim Correlativo As String
Dim NumeroRut As String
Dim FiltroTabVisar As Recordset
Dim NombreArchivo As String
Dim CorrelativoSegunPauta As String
  If CorrelativoSegunLinea = IndiceBotonParaVerFirma Then
    NombreArchivo = UCase(DIRECTORIO_02 + ARCHIVO_EXE09)
    If Dir(NombreArchivo) = Mid$(NombreArchivo, Len(NombreArchivo) - 11) Then
      BtnPauta(CorrelativoSegunLinea).Tag = "0" + Mid$(BtnPauta(CorrelativoSegunLinea).Tag, 2)
      NumeroRut = DesformateaRut(Rut.Caption)
      Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
      Call EjecutarPrograma(NombreArchivo + " " + NumeroRut, 1)
      If SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01) = "OK" Then
        LlamadaProgramaAsociado = True
      End If
    Else
      MsgBox "No se Puede Ejecutar Programa" + Chr(10) + NombreArchivo, 0 + 48, "Falta Programa Asociado"
    End If
  Else
    Visando.GrillaProductos.Col = 4
    Producto = Visando.GrillaProductos.Text
    CorrelativoSegunPauta = Format$(LTrim(Str(CorrelativoSegunLinea)), "00")
    LlamadaProgramaAsociado = False
    NombreArchivo = UCase(DIRECTORIO_02 + Mid$(ARCHIVO_EXE02, 1, Len(ARCHIVO_EXE02) - 9) + _
                                 Producto + CorrelativoSegunPauta + Mid$(ARCHIVO_EXE02, Len(ARCHIVO_EXE02) - 3))
    If Dir(NombreArchivo) = Mid$(NombreArchivo, Len(NombreArchivo) - 11) Then
      NumeroRut = DesformateaRut(Rut.Caption)
      Visando.GrillaProductos.Col = 5
      Correlativo = Visando.GrillaProductos.Text
      Visando.GrillaProductos.Col = 6
      Set PunteroTabVld = BaseVisado.OpenRecordset("SELECT * FROM Tab_Vld WHERE Vld_Prod='" + _
                                         Producto + "' AND Vld_PosPauta=" + CorrelativoSegunPauta + _
                                         " AND Vld_Sistema='VALPTA' ORDER BY Vld_Correl ASC", dbOpenSnapshot)
      If PunteroTabVld.RecordCount > 0 Then
        Do While Not PunteroTabVld.EOF
          If Interpreta(PunteroTabVld("Vld_PosData"), SIMBOLOLARGO) = -1 Then
            Parametro = PunteroTabVld("Vld_PosData")
          Else
            Parametro = Mid$(Visando.GrillaProductos.Text, _
                                          Interpreta(PunteroTabVld("Vld_PosData"), SIMBOLOPOSICION), _
                                          Interpreta(PunteroTabVld("Vld_PosData"), SIMBOLOLARGO))
            Parametro = ObtenerDescripcionGlosa(Producto, Parametro, PunteroTabVld("Vld_Picture"))
          End If
          Retorno = PoneValorINI("Parametros", "Param" + LTrim$(Str(PunteroTabVld("Vld_Correl"))), _
                                                 Parametro, DIRECTORIO_01 + ARCHIVO_INI01)
          PunteroTabVld.MoveNext
        Loop
      Else
        Retorno = PoneValorINI("Parametros", "Param1", NumeroRut, DIRECTORIO_01 + ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "Param2", Producto, DIRECTORIO_01 + ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "Param3", Correlativo, DIRECTORIO_01 + ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "Param4", Format$(LTrim(Str(Tipo)), "0"), DIRECTORIO_01 + _
                                               ARCHIVO_INI01)
        Retorno = PoneValorINI("Parametros", "LlaveVisad", Visando.GrillaProductos.Text, DIRECTORIO_01 + _
                                               ARCHIVO_INI01)
      End If
      PunteroTabVld.Close
      Retorno = PoneValorINI("Parametros", "Retorno", "Ff", DIRECTORIO_01 + ARCHIVO_INI01)
      Call EjecutarPrograma(NombreArchivo, 1)
      If SacaValorINI("Parametros", "Retorno", DIRECTORIO_01 + ARCHIVO_INI01) = "OK" Then
        LlamadaProgramaAsociado = True
      End If
    Else
      MsgBox "No se Puede Ejecutar Programa" + Chr(10) + NombreArchivo, 0 + 48, "Falta Programa Asociado"
    End If
  End If
End Function
Private Function EjecutarActualBoton(ByVal IndiceDependencias As String)
Dim i As Integer
Dim Numero As Integer
  EjecutarActualBoton = True
  If Len(IndiceDependencias) > 0 Then
    Do While InStr(IndiceDependencias, ",") <> 0
      i = InStr(IndiceDependencias, ",")
      Numero = Val(Mid(IndiceDependencias, 1, i - 1))
      IndiceDependencias = Mid$(IndiceDependencias, i + 1)
      If Glosa(Numero).Enabled Then
        EjecutarActualBoton = False
        MsgBox "Se Debe Primero Revisar el Punto " + LTrim$(Str(Numero + 1)) + ".", 0 + 16, "Falta Validación"
      End If
    Loop
    Numero = Val(IndiceDependencias)
    If Glosa(Numero).Enabled Then
      EjecutarActualBoton = False
      MsgBox "Se Debe Primero Revisar el Punto " + LTrim$(Str(Numero + 1)) + ".", 0 + 16, "Falta Validación"
    End If
  End If
End Function
Private Function HabilitarAprobacionVisado() As Boolean
Dim i As Integer
  HabilitarAprobacionVisado = True
  For i = 0 To PantallaVisar.Count - 1
    If TypeOf PantallaVisar.Controls(i) Is CommandButton Then
      If PantallaVisar.Controls(i).Enabled Then
        If Mid$(PantallaVisar.Controls(i).Tag, 1, 1) = "1" Then
          HabilitarAprobacionVisado = False
          Exit For
        End If
      End If
    End If
  Next i
End Function
Private Sub ColocarTickGrillaProductos(ByVal Tick As Image, ByVal ColorFondo As Long)
  With Visando.GrillaProductos
    .Row = .Row
    .Col = 0
    ' FI2026 SHG - CellType (Spread) sin equivalente en MSFlexGrid
    '.CellType = 9
    .BackColor = COLOR_07
    ' FI2026 SHG - TypePictCenter/Stretch/MaintainScale (Spread) sin equivalente en MSFlexGrid; revisar visualmente el ajuste del icono
    '.TypePictCenter = True
    '.TypePictStretch = True
    '.TypePictMaintainScale = True
    ' FI2026_TODO SHG - CellPictureAlignment comentada: no confirmado si existe en esta version de MSFLXGRD.OCX, revisar contra la VM antes de reactivar
    '.CellPictureAlignment = 4 ' FI2026 - flexAlignCenterCenter, reemplaza TypePictCenter/Stretch/MaintainScale
    Set .CellPicture = Tick.Picture ' FI2026 - TypePictPicture (Spread) -> CellPicture (MSFlexGrid)
    For i% = 1 To 3
      .Col = i%
      .BackColor = ColorFondo
    Next i%
  End With
End Sub
Private Sub HabilitarBotonTransmision()
  With Visando
    .DesMarca.Enabled = True
    .EnviarVisado.Enabled = True
    .PortaColor.Visible = True
    .SombraColor.Visible = True
    .GrillaProductos.Row = .GrillaProductos.Row
  End With
End Sub


Public Function recupera_canal() As String

'datos srm para campo rendicion (concilia)
Dim host, apli, mensaje, status2, contexto, funcion, control As String
Dim header, data_srm, data_cliente, convenio, canal12, rendicion As String
Dim ret, ret_srm, largo As Integer

CadenaDatos = ""
canal2 = "008"
rendicion = "01"

data_cliente = "SELECT rend_contratada.rec_cliente FROM rend_contratada "
data_cliente = data_cliente + "WHERE rend_contratada.rec_cliente = " + Mid(GRut_SForma, 1, Len(GRut_SForma) - 1) + " AND "
data_cliente = data_cliente + "rend_contratada.rec_convenio= " + "'" + gCodConvenio + "'" + " AND "
data_cliente = data_cliente + "rend_contratada.rec_rendicion = " + "'" + rendicion + "'" + " AND "
data_cliente = data_cliente + "rend_contratada.rec_canal = '" + canal2 + "'"

header = "b12256781234W00000010npagos    "
mensaje = header + data_cliente

largo = Len(mensaje)
status2 = "00"
funcion = "01"
contexto = "00"
control = ""
host = "U076"
apli = "GS23"
ret = 0

CadenaDatos = Llamada_ConReintentar(host, apli, header, data_cliente, "8K")

'messagebox("Rendicion Contratada",mensaje,information!)

If Mid(CadenaDatos, 14, 4) > 0 Then
   recupera_canal = "S"
Else
   recupera_canal = "N"
End If


End Function
