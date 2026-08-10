Attribute VB_Name = "Modulo1"
'Private ' FI2026 SHG - AisGetUsr deshabilitado (WAIS.DLL 16-bit)
'Declare Function AisGetUsr Lib \"...WAIS.DLL\" As Integer
' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
'Private Declare Function GetModuleUsage Lib "KERNEL32.DLL" (ByVal handler As Long) As Long

Global BaseEnvioSRM As Database
Global BaseVldCta As Database

Global IdUsuario As String
Global CadenaDatos As String
Global GReintento As Boolean
Global Terminar As Boolean
Global IncrementoStatus As Integer
Public Sub EjecutarPrograma(ByVal Programa As String, ByVal Modo As Integer)
' FI2026 - Retorno era Integer: Shell() en VB6 32-bit devuelve el process ID
' (muy por encima de 32767), que desborda un Integer.
Dim Retorno As Long
  Retorno = Shell(Programa, Modo)
  DoEvents
  ' FI2026 - restaura la espera del original 16-bit:
  '     While GetModuleUsage(Retorno) <> 0 : DoEvents : Wend
  ' EjecutarPrograma es SINCRONA: quien la llama asume que el programa
  ' externo ya termino al volver.
  Call WaitProcess(Retorno)
End Sub
Public Function FormateaDatos(NuevaLlave As String, HaydatoS As Boolean, NumVista As String) As Boolean
Dim I As Integer
Dim Data As String
Dim Largo As Integer
Dim Volver As Boolean
Dim Dato_Filtro As String
Dim Posicion As Integer
Dim Continua As String * 1
Dim Desplazar As Integer
Dim Correlativo As Integer
Dim NumeroRut As String
Dim LargoRegistro As Integer
Dim PunteroVld As Recordset
Dim CodigoProducto As String * 3
Dim UbicacionNombre As String
Dim CantidadRegistros As Integer
Dim Tabla As String
Dim UltimoRegistro As String

Tabla = IIf(NumVista = "4086", "tbl_ValMET", "tbl_ValREG")

  Set PunteroVld = BaseVldCta.OpenRecordset(Tabla, dbOpenTable)
  Data = Mid$(CadenaDatos, 8)
  Volver = False
  Continua = Mid$(CadenaDatos, 3, 1)
  Desplazar = 4
  LargoRegistro = Val(Mid(CadenaDatos, 8, 4)) + Desplazar
  CantidadRegistros = Val(Mid(CadenaDatos, 4, 4))
  For I = 1 To CantidadRegistros
     If Val(Mid$(Data, 121 + 4, 5)) > 0 Then
        PunteroVld.AddNew
        PunteroVld!Ctd_Prod = Mid$(Data, 1 + 4, 1)
        PunteroVld!Ctd_Oper = Mid$(Data, 2 + 4, 1)
        PunteroVld!Ctd_Orig = Mid$(Data, 3 + 4, 1)
        PunteroVld!Ctd_Categ = Mid$(Data, 4 + 4, 4)
        PunteroVld!Ctd_Argum = Mid$(Data, 8 + 4, 6)
        PunteroVld!Ctd_Descrip = Mid$(Data, 14 + 4, 50)
        PunteroVld!Ctd_Default = Mid$(Data, 64 + 4, 1)
        PunteroVld!Ctd_Oblig = Mid$(Data, 65 + 4, 1)
        PunteroVld!Ctd_Modif = Mid$(Data, 66 + 4, 1)
        PunteroVld!Ctd_RangoMin = Mid$(Data, 67 + 4, 15)
        PunteroVld!Ctd_RangoMax = Mid$(Data, 82 + 4, 15)
        
        'If InStr(Mid$(Data, 107 + 4, 24), " ") = 0 Then
        
          PunteroVld!Ctd_TipoDat = Mid$(Data, 107 + 4, 14)
          PunteroVld!Ctd_PosVista = Mid$(Data, 121 + 4, 5)
          PunteroVld!Ctd_LargoVta = Mid$(Data, 126 + 4, 5)
          PunteroVld.Update
        'Else
        '  Open "c:\archivo.txt" For Append As #1
        '    Print #1, Mid$(Data, 1, 250)
        '  Close #1
        'End If
     
     End If
     UltimoRegistro = Mid$(Data, 1 + 4, 13)
     Data = Mid$(Data, LargoRegistro + 1)
  Next I
  
  PunteroVld.Close
  
  If Continua = "S" Then
    Volver = True
    NuevaLlave = UltimoRegistro
  End If
  FormateaDatos = Volver
  
End Function
Public Sub Bajar_Datos(Vista As String)

Dim NodoDestino As String
Dim AplicacionDestino As String
Dim Llave As String
Dim HaydatoS As Boolean

  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    Llave = ""
    Usuario = IdUsuario + Vista + "V0000"
    Do
      CadenaDatos = Llamada_ConReintentar(NodoDestino, AplicacionDestino, Usuario, Llave)
      DoEvents
      If Mid$(CadenaDatos, 1, 6) = "000000" Then
          CadenaDatos = Mid$(CadenaDatos, 5)
          
          If Frm_BajCTD.Status.Width + IncrementoStatus < 3570 Then
            Frm_BajCTD.Status.Width = Frm_BajCTD.Status.Width + IncrementoStatus
            Frm_BajCTD.Status.Refresh
          End If
          
          If Not FormateaDatos(Llave, HaydatoS, Vista) Then
             Exit Do
          End If
          Usuario = IdUsuario + Vista + "X0000"
          
      ElseIf Terminar Then
      
        Exit Do
      End If
      
    Loop
  Else
    MsgBox "Vista Nº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If


End Sub
Public Function ObtieneRutUsuarioLogueado() As String
  TIPO_RUT = 0
  ctrlClave.Show 1
  If Not RET_ACCESO Then End
  ObtieneRutUsuarioLogueado = Format$(Mid(RUT_USUARIO, 1, 8), "00000000")
End Function


Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function


Public Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
  
  DeterminaNodo = SINNODO
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - PunteroVistas.Index = PunteroVistas.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroVistas.Seek "=", Vista
  If Not PunteroVistas.NoMatch Then
    If PunteroVistas("Vta_NodoProd") = "U???" Then
        DeterminaNodo = LCase$(Frm_vsd.tx_Nodo.Text)
    Else
        DeterminaNodo = LCase$(PunteroVistas("Vta_NodoProd"))
    End If
    Aplicacion = PunteroVistas("Vta_Aplicacion")
  End If
  PunteroVistas.Close
End Function
