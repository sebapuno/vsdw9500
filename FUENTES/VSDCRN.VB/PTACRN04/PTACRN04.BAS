Attribute VB_Name = "PTACRN04"
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer

Public Function Trae_Nombre(Rut_Cli As String, Rut_Nvo As Variant) As String
  Modo = "L"
  If Environ$("DESA") = "DESA" Then
    NodoDestino = "u005"
  Else
    NodoDestino = "u062"
  End If
  Usuario = IdUsuario + "L001" + Modo
  Llave = Rut_Cli
  CadenaDatos = Llamada(NodoDestino, "CLT0", Usuario, Llave, "1K")
  If Mid$(CadenaDatos, 1, 6) = "000000" Then
     If Mid(Mensaje_Disp, 12, 1) = "K" Then
        Rut_Nvo = Format(Mid(Mensaje_Disp, 4, 8), "0#\.###\.###-") + "K"
     Else
        Rut_Nvo = Format(Mid(Mensaje_Disp, 4, 9), "0#\.###\.###-#")
     End If
     Trae_Nombre = Mid(Mensaje_Disp, 14, 30)
  Else
     Trae_Nombre = "                              "
  End If
End Function


Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(1000, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function

Public Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  Else
    PoneValorINI = False
  End If
End Function

