Attribute VB_Name = "Visado"
' FI2026 SHG - userÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢USER32.DLL: hMenu/iditem/wflags IntegerÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢Long
Private Declare Function DeleteMenu Lib "USER32.DLL" (ByVal hMenu&, ByVal iditem&, ByVal wflags&) As Long
' FI2026 SHG
Private Declare Function GetSystemMenu Lib "USER32.DLL" (ByVal hWnd&, ByVal bRevert&) As Long
' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
'Private ' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
'Declare Function GetModuleUsage Lib \"KERNEL32.DLL\" As Long
Private Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Long

Private Const SC_SIZE = &HF000
Private Const SC_MOVE = &HF010
Private Const MF_BYCOMMAND = &H0

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const DIRECTORIO_02 = "C:\Bin\S\"
Public Const DIRECTORIO_03 = "C:\Data\Cajas\"

Public Const TABLA_DBF01 = "Pag_Not"
Public Const ARCHIVO_INI01 = "Visado.INI"
Public Const ARCHIVO_INI02 = "ErrorSRM.INI"
Public Const ARCHIVO_PIF01 = "Impr.PIF"
Public Const ARCHIVO_IMD01 = "ImVis001.IMD"
Public Const ARCHIVO_TXT01 = "Tran-en2.TXT"
Public Const ARCHIVO_TXT02 = "Ver???.TXT"
Public Const ARCHIVO_EXE01 = "VsdVl???.EXE"
Public Const ARCHIVO_EXE02 = "Pta?????.EXE"
Public Const ARCHIVO_EXE03 = "Exe?????.EXE"
Public Const ARCHIVO_EXE04 = "Srm_Vsd.EXE"
Public Const ARCHIVO_EXE05 = "VsdINot.EXE"
Public Const ARCHIVO_EXE06 = "Rez???00.EXE"
Public Const ARCHIVO_EXE07 = "Rotor???.EXE"
Public Const ARCHIVO_EXE08 = "Ges_???.EXE"
Public Const ARCHIVO_EXE09 = "VerFirma.EXE"
Public Const ARCHIVO_EXE10 = "Amortigd.EXE"
Public Const ARCHIVO_EXE11 = "Frc_???.EXE"

Public Const ARCHIVO_VLCTD = "VLDVSCTD.EXE"
Public Const ARCHIVO_BAJCTD = "BAJCTDLC.EXE"
Public Const ARCHIVO_VISPAR = "VIVLDCTD.EXE"

Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSRM.MDB"
Public Const ARCHIVO_MDB03 = "Oficina.MDB"
Public Const ARCHIVO_MDB04 = "Notarios.MDB"
Public Const ARCHIVO_MDB05 = "Compac?.MDB"

Public Const COLOR_01 = &H55&
Public Const COLOR_02 = &HC0&
Public Const COLOR_03 = &H808000
Public Const COLOR_04 = &HFFFFFF
Public Const COLOR_05 = &HC0FFFF
Public Const COLOR_06 = &HC0C0C0
Public Const COLOR_07 = &H808080
Public Const COLOR_08 = &H80FFFF
Public Const COLOR_09 = &H800080
Public Const COLOR_10 = &H0&

Public Const SINNODO = "Inhabilitado"
Public Const Ejecutable = "Visado"
Public Const NODOMETRO = "NODOM"
Public Const RUTDEFECTO = "12345678"
Public Const SIMBOLOLARGO = "L"
Public Const NODOREGIONAL = "SRM_NS"
Public Const AMBIENTEOFICINA = "OFICINA"
Public Const SIMBOLOPOSICION = "P"
Public Const SIMBOLOLETRACHICA = "<"
Public Const SIMBOLOLETRANORMAL = "="
Public Const AMBIENTEDESARROLLO = "DESA"

Global Oficina As String * 3
Global IdUsuario As String * 8
Global BaseVisado As Database
Global BaseHistory As Database
Global HoraInicioSesion As String * 6
Global DefinicionImpresora As String
Global Version As String

Global Ventana As String
Global GRut_Continua As Boolean
Global GRut_CForma As String
Global GRut_SForma As String
Global GRut_Nombre As String

Global PosicionRut        As String
Global CadenaDatos        As String
Global LargoLlaveOrigen   As Integer
Global PosicionProducto   As String
Global SePresionoTeclaESC As Boolean
Global Filtro             As Boolean
Global Pos_Ftro           As String
Global Dat_Ftro           As String
Global GReintento         As Boolean

Global Logea              As String
Global Vista_Original     As String
Global Usuario_Aux        As String
Global LLave_Aux          As String
Global Vista_Aux          As String
Global HaydatoS           As Boolean
Global Contador           As Integer

'Siebel Fase V
Global EsSiebel           As Boolean
Global Recorte            As Boolean

'Entersoft
'fecha: 10.11.2008
'*****************
Global gCodConvenio       As String
Global gTipoConvenio      As String
Global gTipoRendicion     As String

Public Sub ValidaSiebel()
Dim BaseSiebel      As String
Dim DataConsulta    As String
Dim lsPilotoSiebel  As String
Dim ProcAlm         As String
Dim ResultadoSieb   As String
Dim IntSb As Integer

NodoDestino = SacaValorINI("SYBASE", "NODO", DIRECTORIO_01 & ARCHIVO_INI01)
AplicacionDestino = SacaValorINI("SYBASE", "APLI", DIRECTORIO_01 & ARCHIVO_INI01)
BaseSiebel = SacaValorINI("SYBASE", "BASE", DIRECTORIO_01 & ARCHIVO_INI01)
ProcAlm = SacaValorINI("SYBASE", "PROC", DIRECTORIO_01 & ARCHIVO_INI01)

BaseSiebel = BaseSiebel & Space(10)

DataConsulta = ProcAlm & Space(1) & Environ$("OFICINA") & ",'" & Environ$("SRM_NO") & "'"

IntSb = PoneValorINI("SYBASE", "SIEBEL", "NORMAL", DIRECTORIO_01 & ARCHIVO_INI01)

If Trim$(NodoDestino) <> SINNODO Then
  Usuario = "b12256781234W0000001N" & BaseSiebel
  CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, DataConsulta, "8K")
  If Mid$(CadenaDatos, 1, 6) = "000000" Then
    DataConsulta = Mid(CadenaDatos, 19)
    If InStr(DataConsulta, "~") <> 0 Then
      lsPilotoSiebel = Mid$(DataConsulta, 1, InStr(DataConsulta, "~") - 1)
      'esto implica PILOTO o PRODUCCION
      IntSb = PoneValorINI("SYBASE", "SIEBEL", lsPilotoSiebel, DIRECTORIO_01 & ARCHIVO_INI01)
    End If
  End If
  
End If

End Sub


Public Function EliminaRegistroAntiguo(FchIng As String, DataReg As String, VistaReg As String, OperBorrado As String) As Boolean

Dim AplicacionDestino As String
Dim DiasVigencia As String
Dim FechaMov As String

EliminaRegistroAntiguo = False

FechaMov = FchIng

  If (Date - CDate(FchIng)) > 365 Then
       Usuario = IdUsuario + VistaReg + Trim(OperBorrado)
       DataBorrado = Trim$(VistaReg)
       NodoDestino = DeterminaNodo(AplicacionDestino, VistaReg)
       If Trim$(NodoDestino) <> SINNODO Then
         CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, DataBorrado, "1K")
         If Mid$(CadenaDatos, 1, 6) = "000000" Then
           EliminaRegistroAntiguo = True
         End If
       End If
  End If


End Function

Public Function FormateaFactoring(NuevoUsuario As String, HaydatoS As Boolean, ByVal SinDesplazamiento As Boolean, NVista As String) As Boolean
Dim i             As Integer
Dim Data          As String
Dim Largo         As Integer
Dim Volver        As Boolean
Dim Dato_Filtro   As String
Dim Posicion      As Integer
Dim Continua      As String * 1
Dim Correlativo   As Integer
Dim NumeroRut     As String
Dim PosFin        As Integer
Dim PunteroVsdPrd As Recordset
Dim CodigoProducto    As String * 3
Dim UbicacionNombre   As String
Dim CantidadRegistros As Integer

  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  Data = Mid$(CadenaDatos, 15)
  Volver = False
  Continua = Mid$(CadenaDatos, 3, 1)
  CantidadRegistros = Val(Mid(CadenaDatos, 4, 7))
  For i = 1 To CantidadRegistros  '~
    PosFin = InStr(Data, "~")
    If PosFin > 0 Then
      'NumeroRut = ObtieneNumero(Data, 0, PosicionRut)
      'NumeroRut = NumeroRut + ObtenerDigitoVerificador(NumeroRut)
      NumeroRut = Mid$(Data, 7, 10)
      '*****
      'CodigoProducto = ObtieneNumero(Data, 0, PosicionProducto)
      If NVista = "FC01" Then
         CodigoProducto = "FCT"
      End If
      '****
      If Filtro Then
         Dato_Filtro = ObtieneNumero(Data, 0, Pos_Ftro)
      Else
         Dato_Filtro = ""
      End If
      'CodigoProducto = ObtieneAliasSYBASE(CodigoProducto)
      If ChequeaVigenciaProducto(CodigoProducto, UbicacionNombre, Dato_Filtro) Then
        Correlativo = 0
        Do
          Correlativo = Correlativo + 1
          PunteroVsdPrd.Seek "=", NumeroRut, CodigoProducto, Correlativo
          If PunteroVsdPrd.NoMatch Then
            Exit Do
          End If
        Loop
        PunteroVsdPrd.AddNew
        PunteroVsdPrd("Prd_Data") = Mid$(Data, 1, PosFin - 1)
        PunteroVsdPrd("Prd_Correl") = Correlativo
        PunteroVsdPrd("Prd_RutCli") = NumeroRut
        PunteroVsdPrd("Prd_CodProd") = CodigoProducto
        PunteroVsdPrd.Update
        If (Correlativo = 1) And (Not EsProductoRelacionado(CodigoProducto)) Then
          HaydatoS = True
          Call CrearListaDeClientes(NumeroRut, Mid$(Data, 1, PosFin - 1), UbicacionNombre)
        End If
      End If
    Else
      Continua = "N"
      Exit For
    End If
    Data = Mid$(Data, PosFin + 2) 'en vez de +1)
                                  'porque factoring tiene ~~
  Next i
  PunteroVsdPrd.Close
  If Continua = "S" Then
    Volver = True
    NuevoUsuario = Mid$(NuevoUsuario, 1, 13) + Format$(Val(Mid$(NuevoUsuario, 14, 7)) + CantidadRegistros, "0000000") + Mid$(NuevoUsuario, 21)
  End If
  FormateaFactoring = Volver
End Function
Public Function ChequeaVigenciaProducto(ByVal Codigo As String, Ubicacion As String, Dato_Filtro As String) As Boolean
Dim PunteroTabPrd As Recordset
  Ubicacion = ""
  ChequeaVigenciaProducto = False
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  PunteroTabPrd.Seek "=", Codigo
  If Not PunteroTabPrd.NoMatch Then
    If UCase(PunteroTabPrd("Prd_Tag")) = "SI" Then
      If Filtro Then
         If InStr(RTrim(Dat_Ftro), UCase(Dato_Filtro)) > 0 Then
            Ubicacion = PunteroTabPrd("Prd_Pos_Nom")
            ChequeaVigenciaProducto = True
         End If
      Else
        Ubicacion = PunteroTabPrd("Prd_Pos_Nom")
        ChequeaVigenciaProducto = True
      End If
    End If
  End If
  PunteroTabPrd.Close
End Function


Public Function ObtieneInformacion(ByVal Vista As String, ByVal Operacion As String) As Boolean
Dim Llave As String
Dim Usuario As String
'Dim HaydatoS As Boolean
Dim NodoDestino As String
Dim Desplazamiento As Integer
Dim AplicacionDestino As String * 4
  HaydatoS = False
  NodoDestino = DeterminaNodo(AplicacionDestino, Vista)
  If Trim$(NodoDestino) <> SINNODO Then
    If Mid$(Vista, 1, 3) = "PLA" Then
       Llave = "'" + GRut_SForma + "'"
       Llave = Llave + ",'" + Oficina + "'"
       Usuario = Operacion + " "
    ElseIf Mid$(Vista, 1, 2) = "SY" Then
       Llave = "'" + Oficina + "', NULL, NULL"
       Usuario = Operacion + " "
    ElseIf Mid$(Vista, 1, 2) = "FC" Then
       Llave = Oficina + "000000000000000000000000"
       Usuario = Operacion
    Else
       If Vista = "1664" Then 'para firmas
          Llave = Oficina + "N00000000000000000000000000000000000"
       Else
          Llave = Oficina + "000000000000000000000000000000000000"
       End If
       Usuario = IdUsuario + Vista + Operacion
    End If
    
    Usuario_Aux = Usuario
    LLave_Aux = Llave
    Vista_Aux = Vista
    
    Do
      CadenaDatos = Llamada_ConReintentar(NodoDestino, AplicacionDestino, Usuario, Llave, "8K")
      DoEvents
      If Not SePresionoTeclaESC Then
        If Mid$(CadenaDatos, 1, 6) = "000000" Then
          CadenaDatos = Mid$(CadenaDatos, 5)
          Desplazamiento = False
          If AplicacionDestino = "SPL2" Then
            Desplazamiento = True
          End If
          If Mid$(Vista, 1, 2) = "SY" Then
             If Not FormateaSYBASE(Usuario, HaydatoS, Desplazamiento) Then
               Exit Do
             End If
          ElseIf Mid$(Vista, 1, 2) = "FC" Then
             If Not FormateaFactoring(Usuario, HaydatoS, Desplazamiento, Vista) Then
               Exit Do
             End If
          Else
             If Vista = "1664" And Mid$(CadenaDatos, 15, 1) <> "N" Then 'para firmas
                Exit Do
             End If
             
             If Mid$(Vista, 1, 3) <> "PLA" Then
               If Not FormateaDatos(Llave, HaydatoS, Desplazamiento, Vista) Then
                 Exit Do
               End If
             Else
               Exit Do
             End If
          End If
        Else
          Exit Do
        End If
      Else
        Exit Do
      End If
    Loop
    
  Else
    MsgBox "Vista NÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Âº " + Vista + " No Tiene Nodo Asociado", 0 + 16, "Faltan Datos"
  End If
  ObtieneInformacion = HaydatoS
End Function
Public Function FormateaDatos(NuevaLlave As String, ByRef Hay_Datos As Boolean, ByVal SinDesplazamiento As Boolean, NumVista As String) As Boolean
Dim i As Integer
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
Dim PunteroVsdPrd As Recordset
Dim CodigoProducto As String * 3
Dim UbicacionNombre As String
Dim CantidadRegistros As Integer
  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  Data = Mid$(CadenaDatos, 8)
  Volver = False
  Continua = Mid$(CadenaDatos, 3, 1)
  Desplazar = 4
  LargoRegistro = Val(Mid(CadenaDatos, 8, 4)) + Desplazar
  CantidadRegistros = Val(Mid(CadenaDatos, 4, 4))
  For i = 1 To CantidadRegistros
    If Mid$(Data, 1 + Desplazar + (i - 1) * LargoRegistro, 3) = Oficina Then
      NumeroRut = ObtieneNumero(Data, Desplazar + (i - 1) * LargoRegistro, PosicionRut)
      NumeroRut = NumeroRut + ObtenerDigitoVerificador(NumeroRut)
      CodigoProducto = ObtieneNumero(Data, Desplazar + (i - 1) * LargoRegistro, PosicionProducto)
      
      If Filtro Then
         Dato_Filtro = ObtieneNumero(Data, Desplazar + (i - 1) * LargoRegistro, Pos_Ftro)
      Else
         Dato_Filtro = ""
      End If
      If InStr("00,54,55,56", Environ$("NODOM")) = 0 Then
        If ObtieneAliasRegionalDelProducto(CodigoProducto) <> "" Then
          CodigoProducto = ObtieneAliasRegionalDelProducto(CodigoProducto)
        End If
      End If
      If ChequeaVigenciaProducto(CodigoProducto, UbicacionNombre, Dato_Filtro) Then
        Correlativo = 0
        Do
          Correlativo = Correlativo + 1
          PunteroVsdPrd.Seek "=", NumeroRut, CodigoProducto, Correlativo
          If PunteroVsdPrd.NoMatch Then
            Exit Do
          End If
        Loop
        
        If UCase(Ventana) <> "RUT" Or NumeroRut = GRut_SForma Then
          PunteroVsdPrd.AddNew
          PunteroVsdPrd("Prd_Data") = Mid$(Data, 1 + Desplazar + (i - 1) * LargoRegistro, LargoRegistro)
          PunteroVsdPrd("Prd_Correl") = Correlativo
          PunteroVsdPrd("Prd_RutCli") = NumeroRut
          PunteroVsdPrd("Prd_CodProd") = CodigoProducto
          '***
          PunteroVsdPrd("Prd_VtaOri") = NumVista
          '***
          PunteroVsdPrd.Update
          If (Correlativo = 1) And (Not EsProductoRelacionado(CodigoProducto)) Then
            Hay_Datos = True
            HaydatoS = Hay_Datos
            Call CrearListaDeClientes(NumeroRut, Mid$(Data, 1 + Desplazar + (i - 1) * LargoRegistro, LargoRegistro), _
                                                     UbicacionNombre)
          End If
        End If
      End If
    Else
      Continua = "N"
      Exit For
    End If
    If SinDesplazamiento Then
      Desplazar = 0
    End If
  Next i
  PunteroVsdPrd.Close
  If Continua = "S" Then
    Volver = True
    NuevaLlave = Mid$(Data, 1 + Desplazar + (i - 2) * LargoRegistro, LargoLlaveOrigen)
  End If
  FormateaDatos = Volver
End Function

Public Sub CrearListaDeClientes(ByVal RutCliente As String, ByVal Cadena As String, ByVal Ubicacion As String)
Dim Criterio As String
Dim Usuario As String
Dim NodoDestino As String
Dim PunteroVsdCli As Recordset
Dim NombreCliente As String * 30
Dim AplicacionDestino As String * 4
  Set PunteroVsdCli = BaseVisado.OpenRecordset("Vsd_Cli", dbOpenDynaset)
  Criterio = "Cli_Rut='" + RutCliente + "'"
  NombreCliente = ""
  If Ubicacion = "0,0" Then
    NodoDestino = DeterminaNodo(AplicacionDestino, "1348")
    If Trim$(NodoDestino) <> SINNODO Then
      Usuario = IdUsuario + "1348" + "L"
      CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, Mid$(RutCliente, 1, 9), "8K")
      If Mid$(CadenaDatos, 1, 6) = "000000" Then
        NombreCliente = Mid$(CadenaDatos, 18, 40)
      End If
    End If
  Else
    NombreCliente = Mid$(Cadena, Interpreta(Ubicacion, SIMBOLOPOSICION), _
                                         Interpreta(Ubicacion, SIMBOLOLARGO))
  End If
  NombreCliente = SacarEspaciosIntermedios(NombreCliente)
  PunteroVsdCli.FindFirst Criterio
  If PunteroVsdCli.NoMatch Then
    PunteroVsdCli.AddNew
    PunteroVsdCli("Cli_Rut") = RutCliente
    PunteroVsdCli("Cli_Nombre") = StrConv(NombreCliente, 3)
    PunteroVsdCli.Update
  Else
    If Trim$(PunteroVsdCli("Cli_Nombre")) = "" Then
      PunteroVsdCli.Edit
      PunteroVsdCli("Cli_Nombre") = StrConv(NombreCliente, 3)
      PunteroVsdCli.Update
    End If
  End If
  PunteroVsdCli.Close
End Sub

Public Function ObtieneNumero(ByVal Cadena As String, ByVal Desplazamiento As Integer, DatoBuscado As String) As String
  If InStr(DatoBuscado, ",") = 0 Then
    ObtieneNumero = DatoBuscado
  Else
    Largo = Interpreta(DatoBuscado, SIMBOLOLARGO)
    Posicion = Interpreta(DatoBuscado, SIMBOLOPOSICION)
    ObtieneNumero = Mid$(Cadena, Posicion + Desplazamiento, Largo)
  End If
End Function


Public Function FormateaSYBASE(NuevoUsuario As String, Hay_Datos As Boolean, ByVal SinDesplazamiento As Boolean) As Boolean
Dim i             As Integer
Dim Data          As String
Dim Largo         As Integer
Dim Volver        As Boolean
Dim Dato_Filtro   As String
Dim Posicion      As Integer
Dim Continua      As String * 1
Dim Correlativo   As Integer
Dim NumeroRut     As String
Dim PosFin        As Integer
Dim PunteroVsdPrd As Recordset
Dim CodigoProducto    As String * 3
Dim UbicacionNombre   As String
Dim CantidadRegistros As Integer

  Set PunteroVsdPrd = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
  ' FI2026 - PunteroVsdPrd.Index = PunteroVsdPrd.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  Data = Mid$(CadenaDatos, 15)
  Volver = False
  Continua = Mid$(CadenaDatos, 3, 1)
  CantidadRegistros = Val(Mid(CadenaDatos, 4, 7))
  For i = 1 To CantidadRegistros  '~
    PosFin = InStr(Data, "~")
    If PosFin > 0 Then
      NumeroRut = ObtieneNumero(Data, 0, PosicionRut)
      NumeroRut = NumeroRut + ObtenerDigitoVerificador(NumeroRut)
      CodigoProducto = ObtieneNumero(Data, 0, PosicionProducto)
      
      If Filtro Then
         Dato_Filtro = ObtieneNumero(Data, 0, Pos_Ftro)
      Else
         Dato_Filtro = ""
      End If
      CodigoProducto = ObtieneAliasSYBASE(CodigoProducto)
      If ChequeaVigenciaProducto(CodigoProducto, UbicacionNombre, Dato_Filtro) Then
        Correlativo = 0
        Do
          Correlativo = Correlativo + 1
          PunteroVsdPrd.Seek "=", NumeroRut, CodigoProducto, Correlativo
          If PunteroVsdPrd.NoMatch Then
            Exit Do
          End If
        Loop
        
        '** PARA EXISTIR UNICO REGISTRO EN VISADO
        If Not ExisteTrxVsd(NumeroRut, CodigoProducto, Mid$(Data, 1, PosFin - 1)) Then
          PunteroVsdPrd.AddNew
          PunteroVsdPrd("Prd_Data") = Mid$(Data, 1, PosFin - 1)
          PunteroVsdPrd("Prd_Correl") = Correlativo
          PunteroVsdPrd("Prd_RutCli") = NumeroRut
          PunteroVsdPrd("Prd_CodProd") = CodigoProducto
          PunteroVsdPrd.Update
          If (Correlativo = 1) And (Not EsProductoRelacionado(CodigoProducto)) Then
            Hay_Datos = True
            HaydatoS = Hay_Datos
            Call CrearListaDeClientes(NumeroRut, Mid$(Data, 1, PosFin - 1), UbicacionNombre)
          End If
        End If
      End If
    Else
      Continua = "N"
      Exit For
    End If
    Data = Mid$(Data, PosFin + 1)
  Next i
  PunteroVsdPrd.Close
  If Continua = "S" Then
    Volver = True
    NuevoUsuario = Mid$(NuevoUsuario, 1, 13) + Format$(Val(Mid$(NuevoUsuario, 14, 7)) + CantidadRegistros, "0000000") + Mid$(NuevoUsuario, 21)
  End If
  FormateaSYBASE = Volver
End Function

Public Function ExisteTrxVsd(txRut As String, txCod As String, txDat As String) As Boolean

Dim PunteroVsdPrd As Recordset

txSQL = "Select * From Vsd_Prd Where Prd_CodProd='" + txCod + "' "
txSQL = txSQL + "And Prd_RutCli='" + txRut + "' "
txSQL = txSQL + "And Mid$(Prd_Data,1,60)='" + Mid$(txDat, 1, 60) + "' "

Set PunteroVsdPrd = BaseVisado.OpenRecordset(txSQL, dbOpenDynaset)
If PunteroVsdPrd.RecordCount > 0 Then
  ExisteTrxVsd = True
Else
  ExisteTrxVsd = False
End If
PunteroVsdPrd.Close

End Function


Public Function ObtieneAliasSYBASE(ByVal Producto As String) As String
Dim Criterio As String
Dim PunteroAliasSYBASE As Recordset
  Set PunteroAliasSYBASE = BaseVisado.OpenRecordset("Tab_SybAlias", dbOpenDynaset)
  Criterio = "Syb_Prod_Ori='" + Producto + "'"
  ObtieneAliasSYBASE = Producto
  PunteroAliasSYBASE.FindFirst Criterio
  If Not PunteroAliasSYBASE.NoMatch Then
    ObtieneAliasSYBASE = PunteroAliasSYBASE("Syb_Prod_Nvo")
  End If
  PunteroAliasSYBASE.Close
End Function

Public Function CorrespondeTraer(Modo_Bajada As String)
   CorrespondeTraer = False

   If Modo_Bajada = 0 Then   ' Trae Siempre los datos
      CorrespondeTraer = True
   End If
   
   ' trae si es Regiones
   If Modo_Bajada = 2 And InStr("00,54,55,56", Environ$("NODOM")) = 0 Then
      CorrespondeTraer = True
   End If
   
   ' trae si es Metro
   If Modo_Bajada = 1 And InStr("00,54,55,56", Environ$("NODOM")) <> 0 Then
      CorrespondeTraer = True
   End If
   
End Function

Public Function HayProductosVigentes(DespliegaMensaje As Boolean) As Boolean
Dim Criterio As String
Dim PunteroTabPrd As Recordset
  Criterio = "Prd_Tag='SI'"
  DespliegaMensaje = False
  HayProductosVigentes = False
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenDynaset)
  PunteroTabPrd.FindFirst Criterio
  If Not PunteroTabPrd.NoMatch Then
    DespliegaMensaje = True
    HayProductosVigentes = True
  End If
  PunteroTabPrd.Close
End Function

Public Sub Llena_Historico()
Dim FechaP As String * 8
Dim PunteroVsdHst As Recordset
Dim PunteroHistory As Recordset

Set PunteroHistory = BaseHistory.OpenRecordset("tbl_Hst", dbOpenTable)
Set PunteroVsdHst = BaseVisado.OpenRecordset("vsd_hst", dbOpenTable)

'*** Agregar Nuevos Registros al HistÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³rico
If Not PunteroVsdHst.EOF Then
   PunteroVsdHst.MoveFirst
   While Not PunteroVsdHst.EOF
      PunteroHistory.AddNew
      PunteroHistory!Hst_Rut = PunteroVsdHst!Hst_Rut
      PunteroHistory!Hst_Tipo = PunteroVsdHst!Hst_Tipo
      PunteroHistory!Hst_Data = PunteroVsdHst!Hst_Data
      PunteroHistory!Hst_Fech = PunteroVsdHst!Hst_Fech
      PunteroHistory!Hst_Correl = PunteroVsdHst!Hst_Correl
      PunteroHistory.Update
      PunteroVsdHst.MoveNext
   Wend
End If
PunteroHistory.Close
PunteroVsdHst.Close

'*** Limpia Vsd_Hst
BaseVisado.Execute "DELETE * FROM vsd_hst"

End Sub
Public Function DeterminaNodo(Aplicacion As String, ByVal Vista As String) As String
Dim PunteroVistas As Recordset
Dim BaseEnvioSRM As Database
Dim IniDesa As String

  IniDesa = SacaValorINI("VISADO", "DESA", DIRECTORIO_01 + ARCHIVO_INI01)
  'IniDesa = "DESA"
  DeterminaNodo = SINNODO
  Set BaseEnvioSRM = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB02)
  Set PunteroVistas = BaseEnvioSRM.OpenRecordset("Vistas", dbOpenTable)
  ' FI2026 - PunteroVistas.Index = PunteroVistas.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroVistas.Seek "=", Vista
  If Not PunteroVistas.NoMatch Then
   If Environ$(AMBIENTEDESARROLLO) = "DESA" Or IniDesa = "DESA" Then
      If PunteroVistas("Vta_NodoProd") = "U???" Then
        DeterminaNodo = LCase$(Environ$(NODOREGIONAL))
      Else
        DeterminaNodo = LCase$(PunteroVistas("Vta_NodoDesa"))
      End If
    Else
      If PunteroVistas("Vta_NodoProd") = "U???" Then
        DeterminaNodo = LCase$(Environ$(NODOREGIONAL))
      Else
        DeterminaNodo = LCase$(PunteroVistas("Vta_NodoProd"))
      End If
    End If
    Aplicacion = PunteroVistas("Vta_Aplicacion")
  End If
  PunteroVistas.Close
  BaseEnvioSRM.Close
End Function


Public Function FormateaRut(Rut As String) As String
  If Len(Rut) > 9 Then
    Rut = Mid$(Rut, 2)
  End If
  Rut = String(9 - Len(Rut), "0") + Rut
  FormateaRut = Mid$(Rut, 1, 2) + "." + Mid$(Rut, 3, 3) + "." + Mid$(Rut, 6, 3) + "-" + UCase$(Mid$(Rut, 9, 1))
End Function
Public Function EsProductoRelacionado(ByVal CodigoProducto As String) As Boolean
Dim PunteroTabVisar As Recordset
  EsProductoRelacionado = True
  Set PunteroTabVisar = BaseVisado.OpenRecordset("Tab_Visar", dbOpenTable)
  ' FI2026 - PunteroTabVisar.Index = PunteroTabVisar.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
  PunteroTabVisar.Seek "=", CodigoProducto
  If PunteroTabVisar.NoMatch Then
    EsProductoRelacionado = False
  End If
  PunteroTabVisar.Close
End Function
Public Function GlosaOficina(CodigoOficina As String) As String
Dim Archivo           As String
Dim BaseSucursales    As Database
Dim PunteroSucursales As Recordset
  
  Archivo = UCase("c:\bin\s\baplsgt.exe suc cui")

  If Dir("C:\BIN\S\BAPLSGT.EXE") = "BAPLSGT.EXE" Then
     'Call EjecutarPrograma(Archivo, 1)

     Set BaseSucursales = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
     Set PunteroSucursales = BaseSucursales.OpenRecordset("TblSuc", dbOpenTable)
     ' FI2026 - PunteroSucursales.Index = PunteroSucursales.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
     PunteroSucursales.Seek "=", CodigoOficina
     If PunteroSucursales.NoMatch Then
       GlosaOficina = CodigoOficina + Space(4) + " NO DETERMINADA"
     Else
       GlosaOficina = CodigoOficina + Space(4) + PunteroSucursales!Suc_SucNom
     End If
     PunteroSucursales.Close
     BaseSucursales.Close
  Else
     Set BaseSucursales = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB03)
     Set PunteroSucursales = BaseSucursales.OpenRecordset("Sucursales", dbOpenTable)
     ' FI2026 - PunteroSucursales.Index = PunteroSucursales.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
     PunteroSucursales.Seek "=", CodigoOficina
     If PunteroSucursales.NoMatch Then
       GlosaOficina = CodigoOficina + Space(4) + " NO DETERMINADA"
     Else
       GlosaOficina = CodigoOficina + Space(4) + PunteroSucursales!Suc_Nombre
     End If
     PunteroSucursales.Close
     BaseSucursales.Close
  End If
End Function
Public Function DesformateaRut(ByVal RutCliente As String) As String
Dim Posicion As Integer
Dim NumeroRut As String
  Posicion = 1
  NumeroRut = ""
  While Posicion <= Len(RutCliente)
    If (Mid$(RutCliente, Posicion, 1) <> ".") And (Mid$(RutCliente, Posicion, 1) <> "-") Then
      NumeroRut = NumeroRut + Mid$(RutCliente, Posicion, 1)
    End If
    Posicion = Posicion + 1
  Wend
  NumeroRut = Trim$(NumeroRut)
  DesformateaRut = String(10 - Len(NumeroRut), "0") + NumeroRut
End Function
Function LimpiezaVisado()

Dim Pto_TABVSD   As Recordset
Dim Pto_TABPRD   As Recordset
Dim Pto_VSDPRD   As Recordset
Dim Pto_TABOPER  As Recordset
Dim Pto_VSDCLI As Recordset
Dim VistaOrigen As String

Dim DiasVigencia As String
Dim FechaPrd     As String
Dim AplicacionDestino As String * 4

If WeekDay(Date) = 4 And Environ$("DESA") <> "DESA" Then 'es MiÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â©rcoles
   
  DiasVigencia = SacaValorINI("Visado", "Dias", DIRECTORIO_01 + ARCHIVO_INI01)
  If InStr("15,30,60", DiasVigencia) <> 0 Then
    Set Pto_TABPRD = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
    Set Pto_TABVSD = BaseVisado.OpenRecordset("Tab_Vsd", dbOpenTable)
    Set Pto_VSDPRD = BaseVisado.OpenRecordset("Vsd_Prd", dbOpenTable)
    Set Pto_TABOPER = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
         
    Pto_TABPRD.Index = "Tab_Prd"
    ' FI2026 - Pto_TABVSD.Index = Pto_TABVSD.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
    ' FI2026 - Pto_TABOPER.Index = Pto_TABOPER.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6

    Do While Not Pto_VSDPRD.EOF
      Pto_TABPRD.Seek "=", Pto_VSDPRD("Prd_CodProd")
      Pto_TABOPER.Seek ">", Pto_VSDPRD("Prd_CodProd")
      If Not Pto_TABPRD.NoMatch And Pto_TABOPER("Oper_PROD") = Pto_VSDPRD("Prd_CodProd") Then
        VistaOrigen = Pto_TABPRD("Prd_Vista_Ori")
        FechaPrd = ObtenerDescripcionGlosa(Pto_VSDPRD("Prd_CodProd"), _
                                           Mid$(Pto_VSDPRD("Prd_Data"), _
                                           Interpreta(Pto_TABPRD("Prd_Pos_Fec"), SIMBOLOPOSICION), _
                                           Interpreta(Pto_TABPRD("Prd_Pos_Fec"), SIMBOLOLARGO)), _
                                           Pto_TABPRD("Prd_Pic_Fec"))
        If IsDate(FechaPrd) Then
          If (Date - CDate(FechaPrd)) > Val(DiasVigencia) Then
            Pto_TABVSD.Seek "=", VistaOrigen
            If Not Pto_TABVSD.NoMatch Then
              Usuario = IdUsuario + VistaOrigen + Trim(Pto_TABOPER("Oper_Borrado"))
              DataBorrado = Mid$(Pto_VSDPRD("Prd_Data"), 1, Pto_TABVSD("Vsd_L_LlaveOri"))
              NodoDestino = DeterminaNodo(AplicacionDestino, VistaOrigen)
              If Trim$(NodoDestino) <> SINNODO And Mid$(VistaOrigen, 1, 2) <> "SY" Then
                CadenaDatos = Llamada(NodoDestino, AplicacionDestino, Usuario, DataBorrado, "1K")
                If Mid$(CadenaDatos, 1, 6) = "000000" Then
                  Pto_VSDPRD.Delete
                End If
              End If
            End If
          End If
          
        Else
          '* Fuerza eliminaciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n de archivos mayores a 365 dÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­as
          ' If EliminaRegistroAntiguo(FechaPrd, Pto_VSDPRD!Prd_Data, VistaOrigen, Trim(Pto_TABOPER!Oper_Borrado)) Then
          '    Pto_VSDPRD.Delete
          ' End If
        End If '* IsDate
        
      End If
      Pto_VSDPRD.MoveNext
    Loop
         
    '** agregar borrado de plataformas
         
    '*** eliminar registros de tabla clientes VSD_CLI
    Set Pto_VSDCLI = BaseVisado.OpenRecordset("Select * From VSD_CLI Where " + _
                       "CLI_RUT Not In (Select Distinct Prd_RutCli From VSD_PRD)", dbOpenDynaset)
    Do While Not Pto_VSDCLI.EOF
      Pto_VSDCLI.Delete
      Pto_VSDCLI.MoveFirst
    Loop
    Pto_VSDCLI.Close
    
    '***
         
    Pto_TABVSD.Close
    Pto_TABPRD.Close
    Pto_VSDPRD.Close
    Pto_TABOPER.Close
  End If
End If

End Function

Public Function ObtenerDescripcionGlosa(ByVal CodigoProducto As String, ByVal Datos As String, ByVal Formato As String) As String
Dim Signo As String * 1
Dim Numero As String
Dim TipoFormato As String
Dim IndiceTabGlosa As Recordset
  ObtenerDescripcionGlosa = Datos
  
  'If Mid$(Formato, 1, 2) = "TAL" Then
  
  Select Case Mid$(Formato, 1, 2)
    Case "CS"
      ObtenerDescripcionGlosa = Decodifica(Datos, Interpreta(Mid$(Formato, 3, 7), SIMBOLOPOSICION), _
                                                             Interpreta(Mid$(Formato, 3, 7), SIMBOLOLARGO))
    Case "DC"
      Numero = Decodifica(Trim$(Datos), Interpreta(Mid$(Formato, 3, 7), SIMBOLOPOSICION), _
                                         Interpreta(Mid$(Formato, 3, 7), SIMBOLOLARGO))
      Signo = Mid$(Numero, 1, 1)
      Numero = Mid$(Numero, 2, Len(Numero) - Interpreta(Mid$(Formato, 3, 7), SIMBOLOLARGO) - 1) + "," + _
                       Mid$(Numero, Len(Numero) - Interpreta(Mid$(Formato, 3, 7), SIMBOLOLARGO) + 1)
      While Mid$(Numero, 1, 1) = "0"
        Numero = Mid$(Numero, 2)
      Wend
      ObtenerDescripcionGlosa = Signo + ColocarSeparadorDeMiles(Numero)
    Case "MM"
      Datos = StrConv(Datos, Mid(Formato, 3, 1))
      ObtenerDescripcionGlosa = SacarEspaciosIntermedios(Datos)
    Case "NC"
      ObtenerDescripcionGlosa = Visando.NombreCliente.Caption
    Case "ND"
      Datos = Mid$(Datos, 1, Len(Datos) - Val(Mid(Formato, 3, 3))) + "," + _
                    Mid$(Datos, Len(Datos) - Val(Mid(Formato, 3, 3)) + 1)
      While Mid$(Datos, 1, 1) = "0"
        Datos = Mid$(Datos, 2)
      Wend
      Datos = ColocarSeparadorDeMiles(Datos)
      If Mid$(Formato, 3, 3) = "0" Then
        Datos = Mid$(Datos, 1, InStr(Datos, ",") - 1)
      End If
      ObtenerDescripcionGlosa = Datos
    Case "PF"
      If Mid$(Formato, 3, 8) = "ddmmssaa" Then
        ObtenerDescripcionGlosa = Mid$(Datos, 5, 4) + Mid$(Datos, 3, 2) + Mid$(Datos, 1, 2)
      End If
    Case "RV"
      ObtenerDescripcionGlosa = "0" + IdUsuario + ObtenerDigitoVerificador(IdUsuario)
    Case "DV"
      If Val(Datos) <> 0 Then
         ObtenerDescripcionGlosa = FormateaRut(Datos + ObtenerDigitoVerificador(Datos))
      Else
         ObtenerDescripcionGlosa = ""
      End If
    Case "TG"
      Set IndiceTabGlosa = BaseVisado.OpenRecordset("Tab_Glosa", dbOpenTable)
      ' FI2026 - IndiceTabGlosa.Index = IndiceTabGlosa.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
      IndiceTabGlosa.Seek "=", CodigoProducto, Mid$(Formato, 3, 3), Datos
      If Not IndiceTabGlosa.NoMatch Then
        ObtenerDescripcionGlosa = IndiceTabGlosa("Glosa_Descrip")
      End If
      IndiceTabGlosa.Close
    Case "TX"
      If (Val(Mid$(Datos, 1, 10)) <> 0) And (Mid$(Datos, 11, 1) = "N") Then
        ObtenerDescripcionGlosa = "Abono Cta.Cte en Linea"
      Else
        If Val(Mid$(Datos, 1, 10)) = 0 Then
          ObtenerDescripcionGlosa = "Sin Abono Cta.Cte"
        Else
          ObtenerDescripcionGlosa = "Abono Cta.Cte Normal"
        End If
      End If
    Case "VF"
      ' nota si el producto viene con formato sin el siglo se muestra dd/mm/aa
      ' los productos que son 2000 se muestra dd/mm/aaaa esto es necesario
      ' mantenerlo hasta que los productos que lleguen a visado sean todos 2000
      TipoFormato = Trim$(Mid(Formato, 3, Len(Formato)))
      Select Case TipoFormato
        Case "aaaammdd"
          ObtenerDescripcionGlosa = Mid$(Datos, 7, 2) + "/" + Mid$(Datos, 5, 2) + "/" + Mid$(Datos, 1, 4)
        Case "ddmmssaa", "ddmmaaaa"
          ObtenerDescripcionGlosa = Mid$(Datos, 1, 2) + "/" + Mid$(Datos, 3, 2) + "/" + Mid$(Datos, 5, 4)
      End Select
    Case "TR" 'tipo rendicion
        ObtenerDescripcionGlosa = SacaValorINI("TipoRendicion", Trim(Datos), DIRECTORIO_01 + ARCHIVO_INI01)
    Case "TC" 'Tipo Convenio
        ObtenerDescripcionGlosa = SacaValorINI("TipoConvenio", Trim(Datos), DIRECTORIO_01 + ARCHIVO_INI01)
    Case "CR" 'Cuenta Rendicion
        ObtenerDescripcionGlosa = Trim(Datos)
    Case "CG" 'Contingencia
        ObtenerDescripcionGlosa = SacaValorINI("Contingencia", Trim(Datos), DIRECTORIO_01 + ARCHIVO_INI01)
    
    
    Case Else
         ObtenerDescripcionGlosa = Format$(Datos, Trim$(Formato))
  End Select
End Function
Public Function SacarEspaciosIntermedios(ByVal Nombre As String) As String
Dim i As Integer
Dim EsPrimerBlanco As Boolean
Dim NombreTemporal As String
  NombreTemporal = ""
  If Trim$(Nombre) <> "" Then
    i = 1
    EsPrimerBlanco = False
    While i <= Len(Nombre)
      If Mid$(Nombre, i, 1) <> " " Then
        EsPrimerBlanco = True
        If Mid$(Nombre, i, 1) = "/" Then
           NombreTemporal = NombreTemporal + " "
        Else
           NombreTemporal = NombreTemporal + Mid$(Nombre, i, 1)
        End If
      Else
        If EsPrimerBlanco Then
          EsPrimerBlanco = False
          NombreTemporal = NombreTemporal + " "
        End If
      End If
      i = i + 1
    Wend
  End If
  SacarEspaciosIntermedios = NombreTemporal
End Function
Public Function ObtenerDigitoVerificador(Rut As String) As String
Dim Digito As Integer
Dim Suma As Integer
Dim Caracter As String
  Suma = 0
  For Digito = Len(Rut) To 1 Step -1
    Caracter = Mid$(Rut, Digito, 1)
    If InStr("0123456789", Caracter) > 0 Then
      Suma = Suma + Val(Mid(Rut, Digito, 1)) * ((Len(Rut) - Digito) Mod 6 + 2)
    Else
      Digito = 1
    End If
  Next Digito
  Select Case Suma Mod 11
    Case 0
      ObtenerDigitoVerificador = "0"
    Case 1
      ObtenerDigitoVerificador = "K"
    Case Else
      ObtenerDigitoVerificador = LTrim$(Str(11 - (Suma Mod 11)))
  End Select
End Function
Public Function Decodifica(Numero As String, Posicion As Integer, Largo As Integer) As String
Dim Signo As String * 1
Dim Numerito As String
Dim LargoTotal  As Integer
Dim PosicionInicio As Integer
  If Numero <> "" Then
    If Largo <= 0 Then
      PosicionInicio = 1
      LargoTotal = Posicion - 1
    Else
      PosicionInicio = Posicion + 1
      LargoTotal = Largo - 1
    End If
    If InStr("{ABCDEFGHI", Mid$(Numero, Len(Numero), 1)) <> 0 Then
      Signo = ""
    Else
      Signo = "-"
    End If
    Select Case Mid$(Numero, Len(Numero), 1)
      Case "{", "}"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "0"
      Case "A", "J"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "1"
      Case "B", "K"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "2"
      Case "C", "L"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "3"
      Case "D", "M"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "4"
      Case "E", "N"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "5"
      Case "F", "O"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "6"
      Case "G", "P"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "7"
      Case "H", "Q"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "8"
      Case "I", "R"
        Numerito = Mid$(Numero, PosicionInicio, LargoTotal) + "9"
    End Select
    If Largo <= 0 Then
      Decodifica = Signo + Format$(Numerito, "0000000000000")
    Else
      Decodifica = Signo + Format$(Mid(Numero, 1, Posicion), "0000000000000") + Numerito
    End If
  Else
    Decodifica = " " + String(13, "0")
  End If
End Function
Public Function Interpreta(Cadena As String, DatoBuscado As String) As Integer
Dim PosicionComa As Integer
  PosicionComa = InStr(Cadena, ",")
  If PosicionComa <> 0 Then
    If DatoBuscado = SIMBOLOPOSICION Then
      Interpreta = Val(Mid(Cadena, 1, PosicionComa - 1))
    Else
      Interpreta = Val(Mid(Cadena, PosicionComa + 1, Len(Cadena)))
    End If
  Else
    Interpreta = -1
  End If
End Function
Public Sub CentrarVentana(Forma As Form)
  Forma.Left = (Screen.Width - Forma.Width) / 2
  Forma.Top = (Screen.Height - Forma.Height) / 2
End Sub
Public Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Integer
Dim Ret As Integer
  PoneValorINI = False
  Ret = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If Ret = 1 Then
    PoneValorINI = True
  End If
End Function
Public Function SacaValorINI(Grupo As String, Titulo As String, NombreArchivo As String) As String
Dim ValorGet As Integer
Dim lpReturnedString As String
  lpReturnedString = String$(255, 0)
  nSize% = Len(lpReturnedString)
  ValorGet = GetPrivateProfileString(Grupo, ByVal Titulo, "", lpReturnedString, nSize%, NombreArchivo)
  SacaValorINI = Left$(lpReturnedString$, ValorGet)
End Function
Public Sub FijarVentana(Forma As Form)
Dim hWnd%, hMenu%, Success%
  hWnd% = Forma.hWnd
  hMenu% = GetSystemMenu(hWnd%, 0)
  Success% = DeleteMenu(hMenu%, SC_SIZE, MF_BYCOMMAND)
  Success% = DeleteMenu(hMenu%, SC_MOVE, MF_BYCOMMAND)
End Sub
Public Sub EjecutarPrograma(ByVal Programa As String, ByVal Modo As Integer)
Dim Retorno As Integer
  Retorno = Shell(Programa, Modo)
  DoEvents
  ' FI2026 SHG - GetModuleUsage 16-bit no disponible en Win32
  ' Call WaitProcess(Retorno)  ' requiere modRSrm32.bas

     ' DoEvents  ' FI2026 SHG - orphan from GetModuleUsage
  ' Wend  ' FI2026 SHG - orphan from GetModuleUsage
End Sub
Public Sub BloquearEdicionGrilla(ByVal FilIni As Long, ByVal ColIni As Long, ByVal FilFin As Long, ByVal ColFin As Long, GrillaBloqueada As Object)
  ' FI2026 - BloquearEdicionGrilla migrado de SpreadSheet a MSFlexGrid: Lock/BlockMode no existen,
  ' la grilla es de solo lectura por defecto
End Sub
Public Function ColocarSeparadorDeMiles(ByVal Numero As String) As String
Dim Entero As String
Dim Decimal As String
Dim Cantidad As String
Dim PosicionComa As Integer
  Numero = Trim$(Numero)
  If Mid$(Numero, 1, 1) = "," Then
    Numero = "0" + Numero
  End If
  If Mid$(Numero, Len(Numero), 1) = "," Then
    Numero = Numero + "0"
  End If
  PosicionComa = InStr(Numero, ",")
  Entero = Mid$(Numero, 1, PosicionComa - 1)
  Decimal = Mid$(Numero, PosicionComa + 1)
  Cantidad = ""
  For i% = 1 To Len(Entero)
    Cantidad = Mid$(Numero, Len(Entero) + 1 - i%, 1) + Cantidad
    If i% Mod 3 = 0 Then
      Cantidad = "." + Cantidad
    End If
  Next i%
  If Mid$(Cantidad, 1, 1) = "." Then
    Cantidad = Mid$(Cantidad, 2)
  End If
  ColocarSeparadorDeMiles = Cantidad + "," + Decimal
End Function

Public Sub OrdenarGrilla(ByVal ColumnaInicio As Integer, ByVal ColumnaDeOrdenamiento As Integer, GrillaParaOrdenar As Object, ByVal TipoOrden As Integer)
  ' FI2026 - OrdenarGrilla migrado de SpreadSheet a MSFlexGrid
  Screen.MousePointer = 11
  GrillaParaOrdenar.Row = 1
  GrillaParaOrdenar.Col = ColumnaDeOrdenamiento
  GrillaParaOrdenar.RowSel = GrillaParaOrdenar.Rows - 1
  GrillaParaOrdenar.ColSel = GrillaParaOrdenar.Cols - 1
  If TipoOrden = 1 Then
    GrillaParaOrdenar.Sort = 4  ' flexSortGenericAscending
  Else
    GrillaParaOrdenar.Sort = 5  ' flexSortGenericDescending
  End If
  Screen.MousePointer = 0
End Sub
Public Function ObtieneAliasRegionalDelProducto(ByVal Producto As String) As String
Dim Criterio As String
Dim PunteroTabAlias As Recordset
  Set PunteroTabAlias = BaseVisado.OpenRecordset("Tab_Alias", dbOpenDynaset)
  Criterio = "Alias_Prod_Orig='" + Producto + "'"
  ObtieneAliasRegionalDelProducto = ""
  PunteroTabAlias.FindFirst Criterio
  If Not PunteroTabAlias.NoMatch Then
    ObtieneAliasRegionalDelProducto = PunteroTabAlias("Alias_Prod_Nuev")
  End If
  PunteroTabAlias.Close
End Function

Public Function Corresponde_Oficina(CodOf As String, CodNodom As String) As Boolean

Dim BaseCui    As Database
Dim PunteroCui As Recordset
  
Corresponde_Oficina = True

Set BaseCui = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
Set PunteroCui = BaseCui.OpenRecordset("TblCui", dbOpenTable)
' FI2026 - PunteroCui.Index = PunteroCui.Indexes(0).Name eliminado: Recordset.Indexes esta [restricted] en DAO 3.6
PunteroCui.Seek "=", CodOf
If Not PunteroCui.NoMatch Then
   If InStr("00,54,55,56", CodNodom) <> 0 Then '** Metro
     Corresponde_Oficina = IIf(PunteroCui!cui_cuireg = 13, True, False)
   Else                                        '** Regional
     Corresponde_Oficina = IIf(PunteroCui!cui_cuireg <> 13, True, False)
   End If
End If
PunteroCui.Close
BaseCui.Close

End Function
