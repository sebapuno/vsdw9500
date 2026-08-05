VERSION 4.00
Begin VB.Form Principal 
   Caption         =   "Principal"
   ClientHeight    =   4530
   ClientLeft      =   945
   ClientTop       =   1395
   ClientWidth     =   7275
   Height          =   4875
   Left            =   915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7275
   Top             =   1080
   Visible         =   0   'False
   Width           =   7335
End
Attribute VB_Name = "Principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    ' NO SE REALIZA NINGUNA OPERACION
    
    If Not PoneValorINI("Parametros", "Retorno", "OK", DIRECTORIO_01 + ARCHIVO_INI01) Then
       MsgBox "No Se Puede Actualizar" + Chr(10) + "Archivo " + Chr(34) + "Visado.INI" + Chr(34), 0 + 48, _
              "Error de Escritura"
    End If
    End
End Sub

Private Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Private Function SacaValorINI(ByVal Grupo As String, ByVal Titulo As String, ByVal NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

