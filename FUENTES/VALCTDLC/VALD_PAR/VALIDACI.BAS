Attribute VB_Name = "Validaciones"

Function Vd_CtaBAE(BCuenta As String) As Boolean
  
Dim ponderado As String * 7
Dim d4 As Double
Dim r4 As Double
Dim j As Double
Dim i As Double
Dim r As Double
Dim d As Double
Dim resto_final As Double
Dim total_final As Double
Dim subtotal4 As Double
Dim total As Double
Dim dv As Double
Dim total_total As Double
 
Vd_CtaBAE = False

ponderado = "2212121"
    
If Len(BCuenta) > 11 Then
  BCuenta = Mid$(BCuenta, 2)
End If
    
If Mid$(BCuenta, 4, 1) = "6" Or Mid$(BCuenta, 4, 1) = "7" Then
 
  d4 = Int((Val(Mid$(BCuenta, 4, 1)) * 2) / 10)
  r4 = (Val(Mid$(BCuenta, 4, 1)) * 2) Mod 10
  subtotal4 = d4 + r4
  total_total = 0
  j = 2
  For i = 6 To 11
    d = Int((Val(Mid$(BCuenta, i, 1)) * Val(Mid$(ponderado, j, 1))) / 10)
    r = (Val(Mid$(BCuenta, i, 1)) * Val(Mid$(ponderado, j, 1))) Mod 10
    total = d + r
    total_total = total_total + total
    j = j + 1
  Next i
  
  total_final = total_total + subtotal4
 
  resto_final = total_final Mod 10
 
  If resto_final <> 0 Then
    dv = 10 - resto_final
  Else
    dv = 0
  End If
 
  If dv = Val(Mid$(BCuenta, 5, 1)) Then
    Vd_CtaBAE = True
  End If

End If
 
 
 
End Function

 
Function Vd_CtaBCH(cuenta_ As Variant) As Boolean
 
Dim Cuentav As String

Vd_CtaBCH = False

Cuentav = Strtran(cuenta_, ".", "")
Cuentav = Strtran(cuenta_, "-", "")
Cuentav = UCase(Trim$(cuenta_))
If Len(cuenta_) = 12 Then
   Cuentav = Mid(Cuentav, 2)
End If
 
If Len(Cuentav) = 11 Then
    Ok% = False
    d1% = Val(Mid$(Cuentav, 9, 1)) * 2
    d2% = Val(Mid$(Cuentav, 8, 1)) * 3
    d3% = Val(Mid$(Cuentav, 7, 1)) * 4
    d4% = Val(Mid$(Cuentav, 6, 1)) * 5
    d5% = Val(Mid$(Cuentav, 5, 1)) * 6
    d6% = Val(Mid$(Cuentav, 4, 1)) * 7
    d7% = Val(Mid$(Cuentav, 3, 1)) * 8
    d8% = Val(Mid$(Cuentav, 2, 1)) * 9
    d9% = Val(Mid$(Cuentav, 1, 1)) * 10
 
    mtotal% = d1% + d2% + d3% + d4% + d5% + d6% + d7% + d8% + d9%
 
    If mtotal% = 0 Then
        Ok% = False
    End If
 
    mResto% = mtotal% Mod 11
    dv% = 11 - mResto%
 
    If dv% = 11 Then
        dv% = 0
    End If
 
    If (dv% = Val(Mid$(Cuentav, 10, 2))) Then
        Ok% = True
    End If
 
    Vd_CtaBCH = Ok%
Else
    Vd_CtaBCH = False
End If
 
 
End Function
 

Public Sub Valida_Fecha(FormatoFCH As String, DtFch As String, DtINi As Integer, DtGlosa As String)

Dim FchVd As String

Select Case FormatoFCH
   Case "ddmmyyyy"
      FchVd = Format(DtFch, "&&/&&/&&&&")
      Vd_Fecha = IIf(IsDate(FchVd), True, False)
      
   Case "yyyymmdd"
      FchVd = Mid$(DtFch, 7, 2) + "/" + Mid$(DtFch, 5, 2) + "/" + Mid$(DtFch, 1, 4)
      Vd_Fecha = IIf(IsDate(FchVd), True, False)
      
End Select

If Not Vd_Fecha Then
 Call IngresaError(Str(DtINi), DtFch, DtGlosa, 0)
End If

End Sub

Public Sub Valida_Rut(DatoRut As String, RIni As Integer, RGlosa As String)

If Vd_Rut(Mid$(DatoRut, 1, Len(DatoRut) - 1)) <> Right(DatoRut, 1) Then
   Call IngresaError(Str(RIni), DatoRut, RGlosa, 0)
End If


End Sub

Public Function Vd_Rut(Rut As String) As String
Const Son$ = "1234567890"
Dim A$, b$, dvCal$
Dim i%, aa%, suma%, es%

'limpiar el Rut
For i% = 1 To Len(Rut)
    A$ = Mid$(Rut, i%, 1)
    If InStr(1, Son$, A$) Then b$ = b$ + A$
Next

If Rut <> "" Then
    For i% = 1 To Len(b$)
    A$ = Right$(b$, i%)
    aa% = Val(Left$(A$, 1))

    If i% < 7 Then
       suma% = suma% + aa% * (i% + 1)
    Else
       suma% = suma% + aa% * (i% - 5)
    End If
    Next i%

    es% = 11 - (suma% Mod 11)
    Select Case es%
    Case 11
        dvCal$ = "0"
    Case 10
        dvCal$ = "K"
    Case Else
      dvCal$ = Format$(es%)
    End Select
     
    Vd_Rut = dvCal$
    
End If

End Function
Public Function Vd_Ejec(DatEjec As String, EIni As Integer, EGlosa As String) As Boolean

Vd_Ejec = False


End Function

Public Sub Valida_Cuenta(COrigen As String, CDato As String, CIni As Integer, CGlosa As String)

If Val(CDato) > 0 Then

  Select Case COrigen
    Case "BCH"
      If Not Vd_CtaBCH(CDato) Then
         Vd_Cuenta = False
         Call IngresaError(Str(CIni), CDato, CGlosa, 0)

      End If
    Case "BAE"
      If Not Vd_CtaBAE(CDato) Then
         Vd_Cuenta = False
         Call IngresaError(Str(CIni), CDato, CGlosa, 0)
      End If
  End Select
  
End If

End Sub

Public Sub Valida_Num(Formato As String, DatoNum As String, NGlosa As String, NMax As String, NMin As String)

Dim Numero As String
Dim Entero As String
Dim ParteDecimal As String ' FI2026 - renombrada de 'Decimal': palabra reservada en VB6 32-bit
'*Monto Max y Min   13 Enteros, 2 Decimales  (sin signo)

If Mid$(Formato, 5, 1) = "P" Or InStr("+,-", Mid$(DatoNum, 1, 1)) > 0 Then
  
  '** recorta parte entera del dato
  Entero = Mid$(DatoNum, 1 + 1, Val(Mid$(Formato, 1, 2)))
  '** formatea a largo 13
  Entero = Format$(Val(Entero), String(13, "0"))
  '** coloca signo
  Entero = Mid$(DatoNum, 1, 1) + Entero
  
  If Val(Mid$(Formato, 3, 2)) > 0 Then
    ParteDecimal = Mid$(DatoNum, Val(Mid$(Formato, 1, 2)) + 2, Val(Mid$(Formato, 3, 2)))
  Else
    ParteDecimal = "0"
  End If
  ParteDecimal = Format$(Val(ParteDecimal), String(2, "0"))

  
Else
  Entero = Mid$(DatoNum, 1, Val(Mid$(Formato, 1, 2)))
  Entero = Format$(Val(Entero), String(13, "0"))
      
  ParteDecimal = Mid$(DatoNum, Val(Mid$(Formato, 1, 2)) + 1, Val(Mid$(Formato, 3, 2)))
  ParteDecimal = Format$(Val(ParteDecimal), String(2, "0"))
  
End If

Numero = Entero + ParteDecimal

If Val(NMax) > 0 Or Val(NMin) > 0 Then
  If Val(Numero) > Val(NMax) Or Val(Numero) < Val(NMin) Then
    Call IngresaError(Str(NIni), DatoNum, NGlosa, 0)
  End If
End If

End Sub
 

Public Function Valida_y_Decodifica(Numero As String) As Boolean
Dim Signo As String * 1
Dim Numerito As String
Dim LargoTotal  As Integer

Valida_y_Decodifica = False
Numerito = ""

If Trim$(Numero) <> "" Then
  
  LargoTotal = Len(Numero) - 1
    
  If InStr("{ABCDEFGHI", Mid$(Numero, Len(Numero), 1)) <> 0 Then
    Signo = "+"
  ElseIf InStr("}JKLMNOPQR", Mid$(Numero, Len(Numero), 1)) <> 0 Then
    Signo = "-"
  Else
    Signo = ""
  End If
    
  Select Case Mid$(Numero, Len(Numero), 1)
    Case "{", "}"
        Numerito = Mid$(Numero, 1, LargoTotal) + "0"
    Case "A", "J"
        Numerito = Mid$(Numero, 1, LargoTotal) + "1"
    Case "B", "K"
        Numerito = Mid$(Numero, 1, LargoTotal) + "2"
    Case "C", "L"
        Numerito = Mid$(Numero, 1, LargoTotal) + "3"
    Case "D", "M"
        Numerito = Mid$(Numero, 1, LargoTotal) + "4"
    Case "E", "N"
        Numerito = Mid$(Numero, 1, LargoTotal) + "5"
    Case "F", "O"
        Numerito = Mid$(Numero, 1, LargoTotal) + "6"
    Case "G", "P"
        Numerito = Mid$(Numero, 1, LargoTotal) + "7"
    Case "H", "Q"
        Numerito = Mid$(Numero, 1, LargoTotal) + "8"
    Case "I", "R"
        Numerito = Mid$(Numero, 1, LargoTotal) + "9"
  End Select
  
  If Signo <> "" And IsNumeric(Numerito) Then
    Valida_y_Decodifica = True
  End If

End If

End Function
Public Sub Valida_EjeOrg(ODat As String, OIni As String, OGlosa As String)

Dim PunteroEje As Recordset

If Len(Trim$(ODat)) = 6 Then

  Set PunteroEje = BaseSGT.OpenRecordset("tblejc", dbOpenTable)
  PunteroEje.Index = "ind_tblejc"
  PunteroEje.Seek "=", Mid$(ODat, 1, 3), Mid$(ODat, 4, 3)

  If PunteroEje.EOF Then ' PunteroEje.recordcount
    Call IngresaError(Str(OIni), ODat, OGlosa, 0)
  End If
Else
  Call IngresaError(Str(OIni), ODat, OGlosa, 0)
End If

End Sub

Public Sub Valida_CodCierre(CDato As String, CIni As Integer, CGlosa As String)

If Trim$(CDato) = "" And CodOper = OPER_CIERRE Then
  Call IngresaError(Str(CIni), CDato, CGlosa, 0)
End If

End Sub
