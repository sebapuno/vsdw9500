VERSION 4.00
Begin VB.Form FRMVLD01 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   1755
   ClientTop       =   1425
   ClientWidth     =   6420
   Height          =   4995
   Left            =   1695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6420
   Top             =   1080
   Visible         =   0   'False
   Width           =   6540
End
Attribute VB_Name = "FRMVLD01"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
   FRMVLD01.Visible = False
   Call Ap_Visado
End Sub


