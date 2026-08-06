VERSION 4.00
Begin VB.Form FRMVLD02 
   Caption         =   "Visado de Productos"
   ClientHeight    =   3585
   ClientLeft      =   1395
   ClientTop       =   2085
   ClientWidth     =   6420
   Height          =   3990
   Left            =   1335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   6420
   Top             =   1740
   Width           =   6540
   Begin VB.Frame Frame1 
      Caption         =   "Modalidad de Abono"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2532
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   6168
      Begin VB.OptionButton Option1 
         Caption         =   "Cambiar modalidad a abono normal"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1020
         TabIndex        =   4
         Top             =   1920
         Width           =   4185
      End

      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Abono Automático en línea"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   5
         Top             =   420
         Width           =   2625
      End

      Begin VB.Label Label3 
         Caption         =   $"VLDCRS02.frx":0000
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   555
         Left            =   330
         TabIndex        =   3
         Top             =   990
         Width           =   5550
      End

      Begin VB.Label Label2 
         Caption         =   "Modalidad Actual"
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   8.25
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   330
         TabIndex        =   2
         Top             =   450
         Width           =   1605
      End
   End

   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2520
      TabIndex        =   0
      Top             =   2985
      Width           =   1524
   End
End
Attribute VB_Name = "FRMVLD02"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Salir_Click()

End Sub


Private Sub Command1_Click()
   If Option1.Value Then
      Call AbonoBatch
   End If
      
   p = GrabaParametro("PARAMETROS", "RETORNO", "OK", "C:\DATA\VISADO\VISADO.INI")
   Ap_VsdPrd.Close
   BaseVisado.Close
   End
   
End Sub


