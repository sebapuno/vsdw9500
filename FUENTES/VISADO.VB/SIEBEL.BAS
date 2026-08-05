Attribute VB_Name = "SIEBEL"
Option Explicit
Sub ExtraeHeaderSybase(Data As String, Retorno As String, lHAYMAS As String, lCantidad As Integer)
    Retorno = Left$(Data, 6)
    Data = Mid$(Data, 7)
    If Retorno <> "000000" Then
        lHAYMAS = "N"
        lCantidad = 0
    Else
        lHAYMAS = Left$(Data, 1)
        Data = Mid$(Data, 2)
        lCantidad = Val(Left$(Data, 7))
        Data = Mid$(Data, 12)
    End If
End Sub
Sub ExtraeCampoSP(Data As String, campo As String)
Dim Pos  As Integer
    campo = ""
    Pos = InStr(Data, "~")
    If Pos > 0 Then
        campo = Mid$(Data, 1, Pos - 1)
        Data = Mid$(Data, Pos + 1)
    Else
        campo = Data
        Data = ""
    End If
End Sub
Function ValidaSiebel05(BuscaINI As Boolean) As Boolean
Dim Nodo            As String
Dim Serv            As String
Dim base            As String
Dim datain          As String
Dim dataout         As String
Dim Retorno         As String
Dim lHAYMAS         As String
Dim lCantidad       As Integer
Dim lsPilotoSiebel  As String
Dim fechahoy        As String
Dim I, j            As Integer
Dim Proc            As String

ValidaSiebel05 = False

If BuscaINI Then
  dataout = Trim(SacaValorINI("SIEBEL", "DATA_SIEBEL", DIRECTORIO_01 & ARCHIVO_INI01))
Else
  Nodo = SacaValorINI("SIEBEL", "NODO", DIRECTORIO_01 & ARCHIVO_INI01)
  Serv = SacaValorINI("SIEBEL", "APLI", DIRECTORIO_01 & ARCHIVO_INI01)
  base = SacaValorINI("SIEBEL", "BASE", DIRECTORIO_01 & ARCHIVO_INI01)
  Proc = SacaValorINI("SIEBEL", "PROC5", DIRECTORIO_01 & ARCHIVO_INI01)
  base = base & Space(10)
  datain = Proc & " " & Val(Environ$("oficina")) & ",'" & Environ$("srm_no") & "'"
  dataout = Llamada(Nodo, Serv, "b12256781234W0000001N" & base, datain, "8K")
  I = PoneValorINI("SIEBEL", "DATA_SIEBEL", Trim(dataout), DIRECTORIO_01 & ARCHIVO_INI01)
  fechahoy = Str(CVDate(Format(Now, "dd/mm/yyyy")))
  j = PoneValorINI("SIEBEL", "FECHA", Trim(fechahoy), DIRECTORIO_01 & ARCHIVO_INI01)
End If

If Left(dataout, 6) = "000000" Then
   Call ExtraeHeaderSybase(dataout, Retorno, lHAYMAS, lCantidad)
   Call ExtraeCampoSP(dataout, lsPilotoSiebel)
   'esto implica PILOTO o PRODUCCION
   Select Case UCase(lsPilotoSiebel)
     Case "NORMAL"
        ValidaSiebel05 = False
     Case "SIEBEL"
        ValidaSiebel05 = True
   End Select
End If
End Function




