Attribute VB_Name = "BAS_INI_REMOTO"
Option Explicit

Public Function INI_REMOTO() As String

Dim Usuario As String
Dim LlavePlt As String
Dim NodoDestino As String
Dim AplicacionDestino As String * 4
Dim DataRem As String
Dim AppName As String
Dim KeyName As String
Dim Basename As String
Dim VistaB As String

INI_REMOTO = ""

AppName = SacaValorINI("INIREMOTO", "APLIVSD", DIRECTORIO_01 + ARCHIVO_INI01)
KeyName = SacaValorINI("INIREMOTO", "KEYVSD", DIRECTORIO_01 + ARCHIVO_INI01)
Basename = SacaValorINI("INIREMOTO", "BASE", DIRECTORIO_01 + ARCHIVO_INI01)
VistaB = IIf(InStr(Basename, "_per") > 0, "PLAP", "PLAE")

NodoDestino = DeterminaNodo(AplicacionDestino, VistaB)

Usuario = "b12256781234W0000001N" & Basename & Space(10)
LlavePlt = "sp_lee_ini_remoto '" & AppName & "','" & KeyName & "'"

CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, LlavePlt)
If Mid$(CadenaDatos, 1, 6) = "000000" Then
    If Val(Mid$(CadenaDatos, 8, 7)) > 0 Then
        DataRem = Mid$(CadenaDatos, 19)
        If InStr(DataRem, "~") > 0 Then
            INI_REMOTO = Mid$(DataRem, 1, InStr(DataRem, "~") - 1)
        End If
    End If
End If

End Function
