VERSION 4.00
Begin VB.Form Frm_BajCTD 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizando Tabla de Parámetros"
   ClientHeight    =   960
   ClientLeft      =   4290
   ClientTop       =   1920
   ClientWidth     =   3900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   1365
   Left            =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Top             =   1575
   Width           =   4020
   Begin Threed.SSPanel PanelStatus 
      Height          =   375
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   3765
      _version        =   65536
      _extentx        =   6641
      _extenty        =   661
      _stockprops     =   15
      forecolor       =   128
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      bevelwidth      =   2
      borderwidth     =   1
      bevelinner      =   1
      shadowcolor     =   1
      font3d          =   3
      Begin VB.Label Status 
         BackColor       =   &H00808000&
         Height          =   190
         Left            =   75
         TabIndex        =   2
         Top             =   90
         Width           =   2895
      End
   End

   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Espere un momento"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   675
      TabIndex        =   0
      Top             =   600
      Width           =   2565
   End
End
Attribute VB_Name = "Frm_BajCTD"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Me.Visible = True

Me.Refresh

Set BaseEnvioSRM = OpenDatabase("C:\DATA\VISADO\ENVIOSRM.MDB")
Set BaseVldCta = OpenDatabase("C:\DATA\VISADO\VALCTDLC.MDB")

IdUsuario = ObtieneRutUsuarioLogueado()

BaseVldCta.Execute "Delete * From tbl_ValMET"
BaseVldCta.Execute "Delete * From tbl_ValREG"

Terminar = False

IncrementoStatus = Int(3570 / 64)
Status.Width = 15

PanelStatus.Visible = True
Status.Visible = True

Bajar_Datos ("4086")
Bajar_Datos ("4087")

Status.Width = 3570

BaseEnvioSRM.Close
BaseVldCta.Close

End Sub


