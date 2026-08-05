VERSION 4.00
Begin VB.Form FRMVLD01 
   Caption         =   "Form1"
   ClientHeight    =   4596
   ClientLeft      =   756
   ClientTop       =   1428
   ClientWidth     =   6420
   Height          =   4968
   Left            =   732
   LinkTopic       =   "Form1"
   ScaleHeight     =   4596
   ScaleWidth      =   6420
   Top             =   1080
   Visible         =   0   'False
   Width           =   6468
End
Attribute VB_Name = "FRMVLD01"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
   FRMVLD01.Visible = False
   Call Ap_Visado
End Sub


