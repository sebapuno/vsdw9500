VERSION 4.00
Begin VB.Form FRMVLD03 
   Caption         =   "Visado de Productos"
   ClientHeight    =   3645
   ClientLeft      =   2400
   ClientTop       =   2985
   ClientWidth     =   6420
   Height          =   4050
   Left            =   2340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   6420
   Top             =   2640
   Width           =   6540
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
      Left            =   2376
      TabIndex        =   3
      Top             =   2988
      Width           =   1524
   End

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
      Left            =   108
      TabIndex        =   0
      Top             =   216
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
         Left            =   930
         TabIndex        =   5
         Top             =   1890
         Width           =   4185
      End

      Begin VB.Label Label2 
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
         Left            =   2130
         TabIndex        =   4
         Top             =   450
         Width           =   2625
      End

      Begin VB.Label Label1 
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
         Top             =   480
         Width           =   1665
      End

      Begin VB.Label Label3 
         Caption         =   $"VLDCRS03.frx":0000
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   700
            size            =   12
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   270
         TabIndex        =   1
         Top             =   990
         Width           =   5775
      End
   End
End
Attribute VB_Name = "FRMVLD03"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
   If Option1.Value Then
     Call AbonoBatch
   Else
     If Ap_modo = "1" Then
        Llave = Mid$(Ap_VsdPrd("PRD_DATA"), 1, 490 + 59) + "V"
     End If
     End
   End If
End Sub


