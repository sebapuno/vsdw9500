VERSION 4.00
Begin VB.Form Principal 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visado de Productos"
   ClientHeight    =   6390
   ClientLeft      =   3360
   ClientTop       =   2220
   ClientWidth     =   8985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   7080
   Icon            =   "VISADO.frx":0000
   KeyPreview      =   -1  'True
   Left            =   3300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8985
   Top             =   1590
   Width           =   9105
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel Panel3D1 
      Height          =   6180
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   8835
      _version        =   65536
      _extentx        =   15584
      _extenty        =   10901
      _stockprops     =   15
      forecolor       =   32768
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
      bevelouter      =   1
      bevelinner      =   2
      shadowcolor     =   1
      font3d          =   1
      Begin Threed.SSPanel PortaBotones 
         Height          =   450
         Left            =   90
         TabIndex        =   16
         Top             =   90
         Width           =   3285
         _version        =   65536
         _extentx        =   5794
         _extenty        =   794
         _stockprops     =   15
         forecolor       =   0
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderwidth     =   1
         bevelouter      =   1
         shadowcolor     =   1
         font3d          =   3
         ' FI2026 - VBX.vaBln migrado a VB.Label nativo
         Begin VB.Label EtiquetaAyuda 
            Height          =   375
            Left            =   2775
            Top             =   225
            Width           =   1155
         End

         Begin Threed.SSCommand BtnConsultaPorRut 
            Height          =   345
            Left            =   450
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "Consulta Visado por Rut"
            Top             =   40
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":030A
         End

         Begin Threed.SSCommand BtnReparaArchivos 
            Height          =   345
            Left            =   2415
            TabIndex        =   22
            TabStop         =   0   'False
            Tag             =   "Repara Bases de Datos"
            Top             =   45
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":0758
         End

         Begin Threed.SSCommand BtnSalirVisado 
            Height          =   345
            Left            =   2835
            TabIndex        =   23
            TabStop         =   0   'False
            Tag             =   "Salir Sesión Visado"
            Top             =   45
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":0F6A
         End

         Begin Threed.SSCommand BtnPagoNotario 
            Height          =   345
            Left            =   2040
            TabIndex        =   21
            TabStop         =   0   'False
            Tag             =   "Actualiza Honorarios de Notarios"
            Top             =   45
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":168C
         End

         Begin Threed.SSCommand BtnConfiguraProductos 
            Height          =   345
            Left            =   1665
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "Selecciona Productos a Visar"
            Top             =   45
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":1CB2
         End

         Begin Threed.SSCommand BtnCambiaImpresora 
            Height          =   345
            Left            =   1290
            TabIndex        =   19
            TabStop         =   0   'False
            Tag             =   "Cambia Impresora por Defecto"
            Top             =   45
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":24C4
         End

         Begin Threed.SSCommand BtnImprimeVisados 
            Height          =   345
            Left            =   900
            TabIndex        =   18
            TabStop         =   0   'False
            Tag             =   "Imprime Productos Visados"
            Top             =   45
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":2C22
         End

         Begin Threed.SSCommand BtnConsultaParaVisar 
            Height          =   345
            Left            =   40
            TabIndex        =   17
            TabStop         =   0   'False
            Tag             =   "Consulta Visado Por Oficina"
            Top             =   40
            Width           =   360
            _version        =   65536
            _extentx        =   635
            _extenty        =   609
            _stockprops     =   78
            forecolor       =   8388608
            BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   700
               size            =   8.25
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            bevelwidth      =   1
            roundedcorners  =   0   'False
            picture         =   "VISADO.frx":33BC
         End
      End

      Begin Threed.SSPanel PortaTitulo 
         Height          =   405
         Left            =   1080
         TabIndex        =   14
         Top             =   630
         Width           =   6450
         _version        =   65536
         _extentx        =   11377
         _extenty        =   714
         _stockprops     =   15
         caption         =   "Clientes Por Visar"
         forecolor       =   0
         backcolor       =   -2147483644
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "Arial"
            charset         =   1
            weight          =   700
            size            =   9
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         borderwidth     =   2
         bevelinner      =   2
         shadowcolor     =   1
         font3d          =   3
      End

      Begin Threed.SSPanel PortaGrilla 
         Height          =   3930
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   6450
         _version        =   65536
         _extentx        =   11377
         _extenty        =   6932
         _stockprops     =   15
         caption         =   "Panel3D2"
         forecolor       =   0
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
         bevelinner      =   1
         shadowcolor     =   1
         font3d          =   3
         ' FI2026 - VBX.SpreadSheet migrado a MSFlexGrid (MSFLXGRD.OCX)
         Begin MSFlexGridLib.MSFlexGrid Grilla 
            Height          =   2640
            Left            =   405
            TabIndex        =   15
            Top             =   900
            Width           =   4340
            Cols            =   3
            Rows            =   1
            ScrollBars      =   2  'flexScrollBarVertical
            SelectionMode   =   1  'flexSelectionByRow
            AllowUserResizing=   1  'flexResizeColumns
         End

         Begin Threed.SSPanel IdOficina 
            Height          =   405
            Left            =   495
            TabIndex        =   12
            Top             =   165
            Width           =   5505
            _version        =   65536
            _extentx        =   9710
            _extenty        =   714
            _stockprops     =   15
            caption         =   "Oficina"
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
            bevelouter      =   1
            bevelinner      =   2
            shadowcolor     =   1
            font3d          =   3
         End

         Begin Threed.SSPanel Panel 
            Height          =   2865
            Left            =   4875
            TabIndex        =   3
            Top             =   900
            Width           =   1260
            _version        =   65536
            _extentx        =   2223
            _extenty        =   5054
            _stockprops     =   15
            backcolor       =   -2147483644
            bevelouter      =   0
            Begin Threed.SSFrame Borde 
               Height          =   1665
               Index           =   1
               Left            =   90
               TabIndex        =   4
               Top             =   960
               Width           =   1095
               _version        =   65536
               _extentx        =   1931
               _extenty        =   2937
               _stockprops     =   14
               forecolor       =   0
               alignment       =   2
               font3d          =   3
               shadowstyle     =   1
               Begin Threed.SSCommand BtnCancelar 
                  Height          =   675
                  Left            =   90
                  TabIndex        =   11
                  Top             =   900
                  Width           =   900
                  _version        =   65536
                  _extentx        =   1588
                  _extenty        =   1191
                  _stockprops     =   78
                  caption         =   "&Cancelar"
                  forecolor       =   192
                  font3d          =   1
                  picture         =   "VISADO.frx":3AE2
               End

               Begin Threed.SSCommand BtnProductos 
                  Height          =   675
                  Left            =   90
                  TabIndex        =   10
                  Top             =   180
                  Width           =   900
                  _version        =   65536
                  _extentx        =   1588
                  _extenty        =   1191
                  _stockprops     =   78
                  caption         =   "&Productos"
                  forecolor       =   8388608
                  font3d          =   1
                  picture         =   "VISADO.frx":3FB4
               End
            End

            Begin Threed.SSFrame Borde 
               Height          =   900
               Index           =   0
               Left            =   90
               TabIndex        =   5
               Top             =   60
               Width           =   1095
               _version        =   65536
               _extentx        =   1931
               _extenty        =   1588
               _stockprops     =   14
               caption         =   "Orden"
               font3d          =   3
               shadowstyle     =   1
               Begin Threed.SSOption OpcionNombre 
                  Height          =   210
                  Left            =   120
                  TabIndex        =   7
                  TabStop         =   0   'False
                  Top             =   550
                  Width           =   915
                  _version        =   65536
                  _extentx        =   1614
                  _extenty        =   370
                  _stockprops     =   78
                  caption         =   "&Nombre"
                  forecolor       =   8388736
                  BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
                     name            =   "MS Sans Serif"
                     charset         =   1
                     weight          =   700
                     size            =   8.25
                     underline       =   0   'False
                     italic          =   0   'False
                     strikethrough   =   0   'False
                  EndProperty
                  font3d          =   1
               End

               Begin Threed.SSOption OpcionRut 
                  Height          =   210
                  Left            =   120
                  TabIndex        =   6
                  Top             =   250
                  Width           =   555
                  _version        =   65536
                  _extentx        =   979
                  _extenty        =   370
                  _stockprops     =   78
                  caption         =   "&Rut"
                  forecolor       =   8388736
                  BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
                     name            =   "MS Sans Serif"
                     charset         =   1
                     weight          =   700
                     size            =   8.25
                     underline       =   0   'False
                     italic          =   0   'False
                     strikethrough   =   0   'False
                  EndProperty
                  value           =   -1  'True
                  font3d          =   1
               End
            End
         End

         Begin VB.Label Encabezado 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Nombre Cliente"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   400
               size            =   9.75
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   1625
            TabIndex        =   8
            Top             =   660
            Width           =   2890
         End

         Begin VB.Label Encabezado 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            Caption         =   "Rut Cliente"
            BeginProperty Font 
               name            =   "MS Sans Serif"
               charset         =   1
               weight          =   400
               size            =   9.75
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   405
            TabIndex        =   9
            Top             =   660
            Width           =   1205
         End
      End

      Begin Threed.SSPanel PanelStatus 
         Height          =   375
         Left            =   5940
         TabIndex        =   1
         Top             =   5715
         Visible         =   0   'False
         Width           =   2805
         _version        =   65536
         _extentx        =   4948
         _extenty        =   661
         _stockprops     =   15
         forecolor       =   128
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         bevelwidth      =   2
         borderwidth     =   1
         bevelinner      =   1
         shadowcolor     =   1
         font3d          =   3
         Begin VB.Label Encabezado 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   "[ESC] Cancelar"
            BeginProperty Font 
               name            =   "Arial"
               charset         =   1
               weight          =   400
               size            =   9.75
               underline       =   0   'False
               italic          =   0   'False
               strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   2
            Left            =   765
            TabIndex        =   24
            Top             =   50
            Width           =   1350
         End

         Begin VB.Image Status 
            Height          =   195
            Left            =   90
            Picture         =   "VISADO.frx":4A36
            Stretch         =   -1  'True
            Top             =   90
            Visible         =   0   'False
            Width           =   2630
         End

         Begin VB.Image Avance 
            Height          =   195
            Left            =   90
            Picture         =   "VISADO.frx":5348
            Stretch         =   -1  'True
            Top             =   90
            Visible         =   0   'False
            Width           =   2630
         End
      End

      Begin VB.Label SombraPortaGrilla 
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         Height          =   3930
         Left            =   1140
         TabIndex        =   13
         Top             =   1140
         Width           =   6450
      End

      Begin VB.Image Image1 
         Height          =   630
         Left            =   90
         Picture         =   "VISADO.frx":5C5A
         Top             =   5490
         Width           =   2700
      End

      Begin VB.Image Image2 
         Height          =   1140
         Left            =   3700
         Picture         =   "VISADO.frx":B934
         Top             =   2400
         Visible         =   0   'False
         Width           =   1155
      End
   End

   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   720
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
      helpfile        =   "c:\data\visado\visado.hlp"
      helpkey         =   "F1"
      topage          =   15
   End

   Begin VB.Menu Accion 
      Caption         =   "&Gestión"
      Begin VB.Menu Consultar 
         Caption         =   "&Consultar por Oficina"
         Shortcut        =   ^C
      End

      Begin VB.Menu ConsultaPorRut 
         Caption         =   "Consultar por &Rut"
      End

      Begin VB.Menu Ordenar 
         Caption         =   "&Ordenar"
         Begin VB.Menu RutOrdenado 
            Caption         =   "Por &Rut"
            Shortcut        =   ^R
         End

         Begin VB.Menu NombreOrdenado 
            Caption         =   "Por &Nombre"
            Shortcut        =   ^B
         End
      End

      Begin VB.Menu Linea1 
         Caption         =   "-"
      End

      Begin VB.Menu Impresoras 
         Caption         =   "Con&figurar Impresora"
         Begin VB.Menu TipoImpresora 
            Caption         =   "&Tipo de Impresora..."
         End

         Begin VB.Menu AdministraImpresion 
            Caption         =   "&Administrador de Impresión..."
         End
      End

      Begin VB.Menu Imprimir 
         Caption         =   "&Imprimir"
         Begin VB.Menu ProdVisa 
            Caption         =   "Productos &Visados..."
            Shortcut        =   ^V
         End

         Begin VB.Menu PagoNota 
            Caption         =   "Pagos por &Notario..."
            Shortcut        =   ^N
         End

         Begin VB.Menu FichCred 
            Caption         =   "Ficha de &Créditos"
            Shortcut        =   ^F
         End
      End

      Begin VB.Menu Linea2 
         Caption         =   "-"
      End

      Begin VB.Menu Salir 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End

   Begin VB.Menu Configurar 
      Caption         =   "&Opciones"
      Begin VB.Menu Productos 
         Caption         =   "Seleccionar &Productos..."
         Shortcut        =   ^P
      End

      Begin VB.Menu Oficinas 
         Caption         =   "Seleccionar &Oficina..."
         Shortcut        =   ^O
      End

      Begin VB.Menu Honorarios 
         Caption         =   "Determinar &Honorarios Notario..."
      End

      Begin VB.Menu RFolio 
         Caption         =   "&Rango de Folios..."
      End

      Begin VB.Menu mnu_ConfOfi 
         Caption         =   "&Configura Oficina"
      End

      Begin VB.Menu Linea3 
         Caption         =   "-"
      End

      Begin VB.Menu EnviaLog 
         Caption         =   "&Envia Log Visado (Gestión)"
      End

      Begin VB.Menu mnu_ActParam 
         Caption         =   "Consultar Tabla de Parámetros CTD"
      End

      Begin VB.Menu RepararBD 
         Caption         =   "Reparar Bases de Datos..."
      End

      Begin VB.Menu mnu_ActSGT 
         Caption         =   "Actualizar Tablas SGT"
      End
   End

   Begin VB.Menu Ayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu Contenido 
         Caption         =   "&Contenido"
         Shortcut        =   {F1}
      End

      Begin VB.Menu Referencias 
         Caption         =   "&Referencias..."
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Private ' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer




Private Function BajarDatosVisado(VerMensaje As Boolean) As Boolean
Dim PunteroTabVsd As Recordset
Dim IncrementoStatus As Integer
  If HayProductosVigentes(VerMensaje) Then
    Screen.MousePointer = 11
    Set PunteroTabVsd = BaseVisado.OpenRecordset("Tab_Vsd", dbOpenTable)
    BaseVisado.Execute "DELETE * FROM Vsd_Prd;"
    BaseVisado.Execute "DELETE * FROM Vsd_Cli;"
    BajarDatosVisado = False
    PunteroTabVsd.MoveLast
    IncrementoStatus = Int(2615 / PunteroTabVsd.RecordCount)
    PunteroTabVsd.MoveFirst
    Status.Width = 15
    Status.Visible = True
    PanelStatus.Visible = True
    SePresionoTeclaESC = False
    While (Not PunteroTabVsd.EOF) And (Not SePresionoTeclaESC)
      If CorrespondeTraer(CStr(PunteroTabVsd("Vsd_Bajada"))) Then
         PosicionRut = PunteroTabVsd("Vsd_Pos_Rut")
         LargoLlaveOrigen = PunteroTabVsd("Vsd_L_LLaveOri")
         PosicionProducto = PunteroTabVsd("Vsd_Pos_CodProd")
         Filtro = PunteroTabVsd("Vsd_Filtro")
         If Filtro Then
            Pos_Ftro = PunteroTabVsd("Vsd_Ftro_Pos")
            Dat_Ftro = PunteroTabVsd("Vsd_Ftro_Dat")
         End If
         Vista_Original = PunteroTabVsd("Vsd_Vista_Ori")
         If Mid$(PunteroTabVsd("Vsd_Vista_Ori"), 1, 3) <> "PLA" Then
           If ObtieneInformacion(PunteroTabVsd("Vsd_Vista_Ori"), PunteroTabVsd("Vsd_Operacion")) Then
            BajarDatosVisado = True
           End If
         End If
      End If
      PunteroTabVsd.MoveNext
      Status.Width = Status.Width + IncrementoStatus
    Wend
    PunteroTabVsd.Close
    Status.Width = 2615
    Status.Visible = False
    PanelStatus.Visible = False
    If SePresionoTeclaESC Then
      VerMensaje = False
      BajarDatosVisado = False
    End If
    Call LimpiezaVisado
    Screen.MousePointer = 0
  Else
    MsgBox "Debe Seleccionar Productos a Visar", 0 + 16, "Falta Vigencia Productos"
  End If
End Function

' FI2026 SHG - ObtieneRutUsuarioLogueado reemplazado: AisGetUsr (WAIS.DLL 16-bit)
' reemplazado por login SRM via ctrlClave/modAcceso
Private Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrclave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function

Private Sub AdministraImpresion_Click()
Dim Retorno As Variant
  Screen.MousePointer = 11
  
  If Environ("OS") <> "Windows_NT" Then
    Retorno = Shell("C:\Windows\PrintMan.EXE", 1)
  Else
    'CommonDialog1.FLAGS = &H100000
    'CommonDialog1.ShowPrinter
  End If
    
  Screen.MousePointer = 0
End Sub

Private Sub BtnCambiaImpresora_Click()
  Call TipoImpresora_Click
End Sub

Private Sub BtnCambiaImpresora_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  EtiquetaAyuda.Caption = BtnCambiaImpresora.Tag
End Sub


Private Sub BtnCancelar_Click()
  Call DesplieguePantalla(False)
End Sub
Private Sub BtnCancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call DesplieguePantalla(False)
  End If
End Sub
Private Sub BtnConfiguraProductos_Click()
   BtnConfiguraProductos.Enabled = False
   Call Productos_Click
   BtnConfiguraProductos.Enabled = True
End Sub

Private Sub BtnConfiguraProductos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   EtiquetaAyuda.Caption = BtnConfiguraProductos.Tag
End Sub


Private Sub BtnConsultaParaVisar_Click()
  
'OJO AQUI BORRAR ESTO ES MOMENTANEO
'==============================================
      If Dir(RutaLogLlamadas()) <> "" Then
       If CDate(Format(CDate(Mid(FileDateTime(RutaLogLlamadas()), 1, 10)) + 6, "dd/mm/yyyy")) <= CDate(Format(Now, "dd/mm/yyyy")) Then
        Kill RutaLogLlamadas()
       End If
      End If
'==============================================
      
BtnConsultaParaVisar.Enabled = False
Call Consultar_Click
BtnConsultaParaVisar.Enabled = True

End Sub

Private Sub BtnConsultaParaVisar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 EtiquetaAyuda.Caption = BtnConsultaParaVisar.Tag
End Sub


Private Sub BtnConsultaPorRut_Click()
'OJO AQUI BORRAR ESTO ES MOMENTANEO
'==============================================
      If Dir(RutaLogLlamadas()) <> "" Then
       If CDate(Format(CDate(Mid(FileDateTime(RutaLogLlamadas()), 1, 10)) + 6, "dd/mm/yyyy")) <= CDate(Format(Now, "dd/mm/yyyy")) Then
        Kill RutaLogLlamadas()
       End If
      End If
'==============================================

BtnConsultaPorRut.Enabled = False
Call ConsultaPorRut_Click
BtnConsultaPorRut.Enabled = True

End Sub

Private Sub BtnConsultaPorRut_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
EtiquetaAyuda.Caption = BtnConsultaPorRut.Tag
End Sub


Private Sub BtnImprimeVisados_Click()
  BtnImprimeVisados.Enabled = False
  Call ProdVisa_Click
  BtnImprimeVisados.Enabled = True
End Sub

Private Sub BtnImprimeVisados_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  EtiquetaAyuda.Caption = BtnImprimeVisados.Tag
End Sub


Private Sub BtnPagoNotario_Click()
   Call Honorarios_Click
End Sub

Private Sub BtnPagoNotario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  EtiquetaAyuda.Caption = BtnPagoNotario.Tag
End Sub


Private Sub BtnProductos_Click()

  BtnProductos.Enabled = False
  Call BtnProductos_KeyPress(13)
  BtnProductos.Enabled = True
  
End Sub
Private Sub BtnProductos_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      ' FI2026 - era: X.Row = X.ActiveRow. En Spread, ActiveRow (fila que el
      ' usuario tenia seleccionada) y Row (fila de trabajo que usan .Col/.Text)
      ' eran propiedades DISTINTAS y esa linea copiaba una en la otra. En
      ' MSFlexGrid .Row cumple los dos roles, asi que la asignacion sobra -- y
      ' ademas rompe: .Row vale -1 cuando la grilla no tiene filas de datos, y
      ' asignarle -1 corta con error 30009 (Invalid Row Value).
      ' RowHeight(fila)=0 es el equivalente de RowHidden. Sin fila de datos
      ' seleccionada no hay producto que abrir, asi que no se hace nada.
      If Grilla.Row >= Grilla.FixedRows Then
        If Not Grilla.RowHeight(Grilla.Row) = 0 Then
          Visando.Show
        End If
      End If
    Case 27
      Call BtnCancelar_Click
  End Select
End Sub

Private Sub BtnReparaArchivos_Click()
  RepararArchivos.Show
End Sub

Private Sub BtnReparaArchivos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  EtiquetaAyuda.Caption = BtnReparaArchivos.Tag
End Sub


Private Sub BtnSalirVisado_Click()
  Unload Me
End Sub

Private Sub BtnSalirVisado_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  EtiquetaAyuda.Caption = BtnSalirVisado.Tag
End Sub


Private Sub ConsultaPorRut_Click()

Ventana = "RUT"
GRut_Continua = False

VsdPorRut.Show 1

If GRut_Continua Then
   PortaBotones.Visible = False
   Visando.Show
Else
  If BtnConsultaPorRut.Enabled Then
    BtnConsultaPorRut.SetFocus
  End If
End If


End Sub

Private Sub Consultar_Click()
Dim DigitoRut As String
Dim NumeroRut As String
Dim PunteroVsdCli As Recordset
Dim MostrarMensaje As Boolean
  
Ventana = ""
  
PortaBotones.Enabled = False
Accion.Enabled = False
Configurar.Enabled = False
Ayuda.Enabled = False

If BajarDatosVisado(MostrarMensaje) Then
    Screen.MousePointer = 11
    ' FI2026 - MaxRows=0 (vaciar) -> Rows=1: MSFlexGrid no admite Rows=0 con
    ' FixedRows=1; 1 = solo la fila fija (oculta), sin filas de datos
    Grilla.Rows = 1
    Set PunteroVsdCli = BaseVisado.OpenRecordset("Vsd_Cli", dbOpenTable)
    While Not PunteroVsdCli.EOF
      Grilla.Rows = Grilla.Rows + 1
      Grilla.Row = Grilla.Rows - 1 ' FI2026 - ultima fila de datos (era .MaxRows)
      Grilla.Col = 1
      ' FI2026 - Spread-only property, no MSFlexGrid equivalent
      ' FI2026 - Spread-only property, no MSFlexGrid equivalent
      ' FI2026 - Spread-only property, no MSFlexGrid equivalent
      DigitoRut = Mid$(PunteroVsdCli("Cli_Rut"), Len(PunteroVsdCli("Cli_Rut")))
      NumeroRut = Mid$(PunteroVsdCli("Cli_Rut"), 1, Len(PunteroVsdCli("Cli_Rut")) - 1)
      NumeroRut = Format$(NumeroRut, "###,###,###") + "-" + DigitoRut
      Grilla.Text = String(13 - Len(NumeroRut), " ") + NumeroRut
      Grilla.Col = 2
      Grilla.Text = PunteroVsdCli("Cli_Nombre")
      For j% = 1 To Grilla.Cols - 1 ' FI2026 - era .MaxCols
        Grilla.Col = j%
        Grilla.BackColor = COLOR_06
        ' FI2026 - Spread-only property, no MSFlexGrid equivalent
        If j% Mod 2 <> 0 Then
          ' FI2026 - Spread-only property, no MSFlexGrid equivalent
        End If
        ' FI2026 - Spread-only property, no MSFlexGrid equivalent
      Next j%
      PunteroVsdCli.MoveNext
    Wend
    PunteroVsdCli.Close
    Screen.MousePointer = 0
    Call OrdenarGrilla(1, 1, Grilla, 1)
    Call DesplieguePantalla(True)
Else
    If MostrarMensaje Then
      MsgBox "No se ha Localizado Información a Visar", 0 + 16, "Datos No Recepcionados"
    End If
End If

Accion.Enabled = True
Configurar.Enabled = True
Ayuda.Enabled = True
PortaBotones.Enabled = True

End Sub
Private Sub Contenido_Click()

Call Form_KeyUp(vbKeyF1, 0)

End Sub

Private Sub Encabezado_Click(Index As Integer)
  Select Case Index
    Case 0
      OpcionRut.Value = True
    Case 1
      OpcionNombre.Value = True
  End Select
End Sub

Private Sub EnviaLog_Click()
   Archivo = UCase("c:\bin\s\enviavsd.bat")
   Call EjecutarPrograma(Archivo, 1)
End Sub

Private Sub FichCred_Click()

If Dir("c:\data\visado\hist_vsd.mdb") <> "HIST_VSD.MDB" Then
     MsgBox "No se encontró la base de datos Hist_vsd.mdb", 16, "No se pueden consultar datos"
     Screen.MousePointer = 0
     Exit Sub
Else
  ImprimeFichaCredito.Show
End If

End Sub
Private Sub Form_Activate()
Dim NumeroFilaDestino As Integer
  If Grilla.Visible Then
    ' FI2026 - RowHidden no existe en MSFlexGrid; RowHeight(fila)=0 es el equivalente
    If Grilla.RowHeight(Grilla.Row) = 0 Then
      ' FI2026 - .Rows - 1: era .MaxRows (ultima fila de datos)
      If Grilla.Row = Grilla.Rows - 1 Then
        NumeroFilaDestino = 1
      Else
        NumeroFilaDestino = Grilla.Row + 1
      End If
      Grilla.Row = NumeroFilaDestino
    End If
  End If
  Screen.MousePointer = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    SePresionoTeclaESC = True
  End If
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF1 Then
  CommonDialog1.HelpFile = "c:\data\visado\visado.hlp"
  CommonDialog1.HelpCommand = &H101
  CommonDialog1.HelpKey = "principal"
  CommonDialog1.ShowHelp
End If


End Sub


Private Sub Form_Load()
Dim Archivo As String
Dim Respuesta As Integer
Dim TablasSgt As String
Dim fecha_ini     As Date
Dim BuscaINI      As Boolean
Dim DataSiebel    As String

If Not App.PrevInstance Then
    Screen.MousePointer = 11
    
    Version = SacaValorINI("Visado", "Version", DIRECTORIO_01 + ARCHIVO_INI01)
    Logea = SacaValorINI("Logea", "Flag", DIRECTORIO_01 + ARCHIVO_INI01)
    
    
    ' FI2026 - carga de los parametros de acceso a la base (SendQuery del login).
    ' Estaban sin cargar: modAcceso.bas tenia Inicializar_Acceso COMENTADA porque
    ' llamaba a SacaValorDeINI, helper que no existe en este proyecto (el typo es
    ' del original: la funcion real se llama SacaValorINI, sin el "De", y es
    ' identica linea por linea). Sin esto NODO_BD queda vacio, el DLL no resuelve
    ' el nodo y sale al 127.0.0.1 -> SRM error 41 -> el login nunca autentica y
    ' ObtieneRutUsuarioLogueado hace End: la aplicacion moria en silencio.
    ' Se replica el patron de WCRE9500 (MDICRED.FRM), que lee de [BASE_DATOS].
    NAME_BD = SacaValorINI("BASE_DATOS", "BDATOS", DIRECTORIO_01 + ARCHIVO_INI01)
    SERV_BD = SacaValorINI("BASE_DATOS", "APLI", DIRECTORIO_01 + ARCHIVO_INI01)
    NODO_BD = SacaValorINI("BASE_DATOS", "NODO", DIRECTORIO_01 + ARCHIVO_INI01)
    If Trim$(NAME_BD) = "" Or Trim$(SERV_BD) = "" Or Trim$(NODO_BD) = "" Then
      MsgBox "Faltan parametros [BASE_DATOS] (NODO/APLI/BDATOS) en " & ARCHIVO_INI01, 16, "Control de Acceso"
      End
    End If
    If IsNumeric(SacaValorINI("FUNCIONALIDAD", "FUNCIONALIDAD", DIRECTORIO_01 + ARCHIVO_INI01)) Then
      FUNCIONALIDAD = CInt(SacaValorINI("FUNCIONALIDAD", "FUNCIONALIDAD", DIRECTORIO_01 + ARCHIVO_INI01))
    End If
    IdUsuario = ObtieneRutUsuarioLogueado()
    HoraInicioSesion = Format$(Now, "hhmmss")
   
    '** validar estacion SIEBEL
    Call ValidaSiebel
    
    'jsa 30/03/2007
    'Recorte de Plataforma Siebel Fase V
    'Recorte de Visado de Cuentas Vista
    'Verifica si estación esta en Siebel V
    '***********************************************
    BuscaINI = False
    fecha_ini = CVDate(SacaValorINI("SIEBEL", "FECHA", DIRECTORIO_01 + ARCHIVO_INI01))
    DataSiebel = Trim(SacaValorINI("SIEBEL", "DATA_SIEBEL", DIRECTORIO_01 + ARCHIVO_INI01))
    If fecha_ini < CVDate(Format(Now, "dd/mm/yyyy")) Then
       BuscaINI = False
    Else
       If DataSiebel <> "" Then
          BuscaINI = True
       End If
    End If
    
    EsSiebel = False
    If ValidaSiebel05(BuscaINI) Then
      EsSiebel = True
    End If
    '********************************************************
    
    
    If Es_Primera_Vez Then
    
      'OJO AQUI BORRAR ESTO ES MOMENTANEO
      '==============================================
      'If Dir("d:\VSD_LOG.TXT") = "VSD_LOG.TXT" Then
      ' If CDate(Format(CDate(Mid(FileDateTime("d:\VSD_LOG.TXT"), 1, 10)) + 6, "dd/mm/yyyy")) <= CDate(Format(Now, "dd/mm/yyyy")) Then
      '  Kill "d:\VSD_LOG.TXT"
      ' End If
      'End If
      '==============================================
      
      TablasSgt = SacaValorINI("Visado", "TablasSGT", DIRECTORIO_01 + ARCHIVO_INI01)
      
      Archivo = UCase("c:\bin\s\baplsgt.exe " + TablasSgt) '"suc cui aec")
      If Dir("C:\BIN\S\BAPLSGT.EXE") = "BAPLSGT.EXE" Then
        Call EjecutarPrograma(Archivo, 1)
      End If
      
      Archivo = UCase(DIRECTORIO_02 + ARCHIVO_BAJCTD)
      If Dir(Archivo) = Mid$(Archivo, Len(Archivo) - 11) Then
        Call EjecutarPrograma(Archivo, 1)
      End If
            
    End If
    
    Set BaseVisado = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB01)
    Call CentrarVentana(Me)
    Call FijarVentana(Me)
    Call DesplieguePantalla(False)
    Oficina = SacaValorINI("Visado", "Oficina", DIRECTORIO_01 + ARCHIVO_INI01)
    If Trim$(Oficina) = "" Then
      Oficina = Environ$(AMBIENTEOFICINA)
      Respuesta = PoneValorINI("Visado", "Oficina", Oficina, DIRECTORIO_01 + ARCHIVO_INI01)
    End If
    
    If Not IsNumeric(Oficina) Then
       MsgBox "La variable Oficina no ha sido configurada, las consultas presentarán problemas." + _
          Chr(10) + "Utilice el Configurador de Posventa para solucionar esto.", 16, "Visado de Productos"
    End If
    
    DefinicionImpresora = SacaValorINI("Visado", "Impresora", DIRECTORIO_01 + ARCHIVO_INI01)
    If InStr("00,54,55,56", Environ$("NODOM")) = 0 Then
       Me.Caption = "Visado de Productos " + Version + " - Regional"
    Else
       Me.Caption = "Visado de Productos " + Version + " - Metro"
    End If
    Panel3D1.Width = Screen.Width - 180
    Panel3D1.Height = Screen.Height - 750
    Image1.Top = Panel3D1.Height - Image1.Height - 90
    PortaGrilla.Left = (Panel3D1.Width - PortaGrilla.Width) / 2
    PortaGrilla.Top = (Panel3D1.Height - PortaGrilla.Height) / 2
    PortaTitulo.Left = (Panel3D1.Width - PortaTitulo.Width) / 2
    PortaTitulo.Top = PortaGrilla.Top - PortaTitulo.Height - 30
    PanelStatus.Left = Panel3D1.Width - PanelStatus.Width - 90
    PanelStatus.Top = Panel3D1.Height - PanelStatus.Height - 90
    SombraPortaGrilla.Left = PortaGrilla.Left + 60
    SombraPortaGrilla.Top = PortaGrilla.Top + 60
    ' FI2026 - ColWidth: Spread mide en CARACTERES, MSFlexGrid en TWIPS
    ' (~120 twips por caracter). Original: ColWidth(1)=10 y ColWidth(2)=24.
    ' Los indices tambien se corrigieron: una migracion previa los habia puesto
    ' en 0 y 1, pero el codigo escribe el Rut en .Col = 1 y el Nombre en .Col = 2,
    ' asi que los anchos se estaban aplicando a las columnas equivocadas.
    Grilla.ColWidth(0) = 0      ' columna fija (el Spread no mostraba encabezados)
    Grilla.ColWidth(1) = 1200   ' Rut    (era 10 caracteres)
    Grilla.ColWidth(2) = 3000   ' Nombre (era 24 caracteres)
    Grilla.RowHeight(0) = 0     ' fila fija oculta (DisplayColHeaders=0)
    Call ConfigurarNodoDestinoSAT
Else
    MsgBox "Visado Ya Está Ejecutándose", 0 + 16, "Doble Ejecución"
    End
End If
  
App.HelpFile = "c:\data\visado\visado.hlp"
 'CommonDialog1.HelpKey = "principal"
If Environ("OS") = "Windows_NT" Then
   AdministraImpresion.Visible = False
Else
   AdministraImpresion.Visible = True
End If
   
End Sub

Private Sub Form_Resize()

  If Me.WindowState <> 1 Then
    Me.WindowState = 2
  End If
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
  
  On Error GoTo error_base
  
  '* se cambia de 120 a 365 días el histórico
  FechaPasada = Format$(Year(Date - 365), "0000") + Format$(Month(Date - 365), "00") + _
              Format$(Day(Date - 365), "00")
              
  Set BaseHistory = OpenDatabase(DIRECTORIO_01 + "Hist_vsd.mdb")
  BaseHistory.Execute "DELETE * FROM tbl_Hst WHERE Val(Hst_Fech)<" + FechaPasada + ";"
   
  Llena_Historico
  
  BaseHistory.Close
  DBEngine.CompactDatabase "C:\data\visado\hist_vsd.mdb", "C:\data\visado\hist_tmp.mdb"
  If Dir("C:\data\visado\hist_tmp.mdb") = "HIST_TMP.MDB" Then
     Kill "c:\data\visado\hist_vsd.mdb"
     FileCopy "C:\data\visado\hist_tmp.mdb", "C:\data\visado\hist_vsd.mdb"
     Kill "c:\data\visado\hist_tmp.mdb"
     Kill "c:\data\visado\hist_tmp.ldb"
  End If
  
'  BaseVisado.Close
  Exit Sub

error_base:
   Resume Next
  
  
End Sub
Private Sub Grilla_DblClick()
  Call BtnProductos_KeyPress(13)
End Sub

Private Sub Grilla_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case 13
      Call BtnProductos_KeyPress(13)
    Case 27
      Call BtnCancelar_Click
  End Select
End Sub
Private Sub Honorarios_Click()
   ConfiguraHonorarios.Show
End Sub

Private Sub mnu_ActParam_Click()

Dim Archivo As String

Screen.MousePointer = 11
Principal.Enabled = False
    
Archivo = UCase(DIRECTORIO_02 + ARCHIVO_VISPAR)
If Dir(Archivo) = Mid$(Archivo, Len(Archivo) - 11) Then
  Call EjecutarPrograma(Archivo, 1)
Else
  MsgBox "No se encuentra el archivo " + Archivo, 16, "Operación Cancelada"
End If

Principal.Enabled = True
Screen.MousePointer = 0

End Sub

Private Sub mnu_ActSGT_Click()

Dim Archivo As String
Dim Respuesta As Integer
Dim TablasSgt As String

Screen.MousePointer = 11
TablasSgt = SacaValorINI("Visado", "TablasSGT", DIRECTORIO_01 + ARCHIVO_INI01)
    
Archivo = UCase("c:\bin\s\baplsgt.exe ")
If Dir("C:\BIN\S\BAPLSGT.EXE") = "BAPLSGT.EXE" Then
   Call EjecutarPrograma(Archivo + TablasSgt, 1) '"suc cui")
Else
   MsgBox "No se encuentra el archivo " + Archivo, 16, "Operación Cancelada"
End If

Screen.MousePointer = 0

End Sub

Private Sub mnu_ConfOfi_Click()

Dim BaseSucursales    As Database
Dim PunteroSucursales As Recordset
Dim HaySuc As Boolean

HaySuc = False

Set BaseSucursales = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
Set PunteroSucursales = BaseSucursales.OpenRecordset("Select Distinct SUC_SUCCOD,SUC_SUCNOM from TblSuc", dbOpenDynaset)
If PunteroSucursales.RecordCount > 0 Then
   HaySuc = True
End If
PunteroSucursales.Close
BaseSucursales.Close

If HaySuc Then
  GrupoOficina.Show
Else
   MsgBox "No existen datos de sucursales en tablas SGT.", 48, "No se puede configurar oficina"
End If

End Sub

Private Sub NombreOrdenado_Click()
  OpcionNombre.Value = True
End Sub
Private Sub Oficinas_Click()
   ConfiguraOficina.Show
End Sub
Private Sub OpcionNombre_Click(Value As Integer)
  Call OrdenarGrilla(1, 2, Grilla, 1)
End Sub
Private Sub OpcionNombre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    BtnCancelar_Click
  End If
End Sub
Private Sub OpcionRut_Click(Value As Integer)
  Call OrdenarGrilla(1, 1, Grilla, 1)
End Sub
Private Sub OpcionRut_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    BtnCancelar_Click
  End If
End Sub

Private Sub PagoNota_Click()
  ImprimePagoNotario.Show
End Sub

Private Sub Productos_Click()
   ConfiguraProducto.Show
End Sub
Private Sub ProdVisa_Click()

If Dir("c:\data\visado\hist_vsd.mdb") <> "HIST_VSD.MDB" Then
     MsgBox "No se encontró la base de datos Hist_vsd.mdb", 16, "No se pueden consultar datos"
     Screen.MousePointer = 0
     Exit Sub
Else
   '  ImprimeProductoVisado.Show
     ImprimeVisadosNew.Show
End If

End Sub

Private Sub Referencias_Click()
  ReferenciasVisado.Show
End Sub
Private Sub RepararBD_Click()
  RepararArchivos.Show
End Sub

Private Sub RFolio_Click()
   Archivo = UCase("c:\bin\s\vsdfolio.exe")
   Call EjecutarPrograma(Archivo, 1)
End Sub

Private Sub RutOrdenado_Click()
  OpcionRut.Value = True
End Sub
Private Sub Salir_Click()
  Unload Me
End Sub
Private Sub DesplieguePantalla(ByVal Estado As Boolean)
  Imprimir.Enabled = Not (Estado)
  Consultar.Enabled = Not (Estado)
  Configurar.Enabled = Not (Estado)
  Impresoras.Enabled = Not (Estado)
  Panel.Visible = Estado
  Ordenar.Enabled = Estado
  PortaGrilla.Visible = Estado
  PortaTitulo.Visible = Estado
  PortaBotones.Visible = Not (Estado)
  PortaBotones.Enabled = Not (Estado)
  SombraPortaGrilla.Visible = Estado
  If Estado Then
    IdOficina.Caption = GlosaOficina(Oficina)
    Grilla.SetFocus
  End If
End Sub



Private Sub TipoImpresora_Click()
  Screen.MousePointer = 11
  DefineImpresora.Show
End Sub

Private Sub ConfigurarNodoDestinoSAT()
Dim ArchivoSRM As Database
Dim PunteroSAT As Recordset
Dim PunteroVistas As Recordset
  If Not Environ$(AMBIENTEDESARROLLO) = "DESA" Then
     Set ArchivoSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
     Set PunteroSAT = ArchivoSRM.OpenRecordset("Sat")
     Set PunteroVistas = ArchivoSRM.OpenRecordset("Vistas")
     PunteroSAT.Index = "Srm_Ns"
     PunteroSAT.Seek "=", UCase(Environ(NODOREGIONAL))
     If Not PunteroSAT.NoMatch Then
       PunteroVistas.Index = "Vistas"
       PunteroVistas.Seek "=", "SAT_"
       If PunteroVistas("Vta_NodoProd") <> PunteroSAT("Sat_Nodo_Sat") Then
         PunteroVistas.Edit
         PunteroVistas("Vta_NodoProd") = PunteroSAT("Sat_Nodo_Sat")
         PunteroVistas.Update
       End If
     Else
       MsgBox "La Variable de Ambiente " + Chr(34) + "SRM_NS" + Chr(34) + " No Fue" + Chr(10) + _
              "Encontrada en Tabla " + Chr(34) + "SAT" + Chr(34) + ", las Transacciones" + Chr(10) + _
              "Via S.A.T. Podrían Tener Conflictos", 0 + 48, "Advertencia"
     End If
     PunteroVistas.Close
     PunteroSAT.Close
     ArchivoSRM.Close
  End If
End Sub


Public Function Es_Primera_Vez() As Boolean

Dim FchAct As String
Dim Respuesta As Integer

Es_Primera_Vez = False
FchAct = SacaValorINI("Visado", "FECHA", DIRECTORIO_01 + ARCHIVO_INI01)
    
If FchAct = "" Then
   FchAct = Format(DateAdd("d", -1, Date), "dd/mm/yyyy")
End If
    
If CDate(FchAct) < Date Then
  Es_Primera_Vez = True
  Respuesta = PoneValorINI("Visado", "FECHA", Format(Date, "dd/mm/yyyy"), DIRECTORIO_01 + ARCHIVO_INI01)
End If


End Function

