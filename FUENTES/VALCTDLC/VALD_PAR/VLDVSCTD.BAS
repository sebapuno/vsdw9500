Attribute VB_Name = "ExeVSCTD00"
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal X%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer

Global Const SWP_NOSIZENOMOVE = 3
Global Const HWND_TOPMOST = -1

Public Const DIRECTORIO_01 = "C:\Data\Visado\"
Public Const DIRECTORIO_02 = "C:\Data\AdmSGT\"

Public Const ARCHIVO_INI01 = "Visado.INI"

Public Const ARCHIVO_MDB01 = "Visado.MDB"
Public Const ARCHIVO_MDB02 = "EnvioSRM.MDB"
Public Const ARCHIVO_MDB03 = "ValCTDLC.MDB"
Public Const ARCHIVO_MDB04 = "TABSGT.MDB"

Public Const SIMBOLOLARGO = "L"
Public Const SIMBOLOPOSICION = "P"

Public Const OPER_APERTURA = "0"
Public Const OPER_CIERRE = "1"
Public Const OPER_RENOVACION = "2"
Public Const OPER_MODIFICACION = "3"

Global BaseVisado As Database
Global BaseSGT As Database
Global BaseVLDCTD As Database

Type Estructura_Error
  Tipo_Error As Integer  ' 0 Normal - 1 Tabla
  Posicion  As String
  Dato  As String
  Descripcion As String
End Type
  
Global Errores() As Estructura_Error

Global CtdLcd As String
Global Origen As String
Global TipoProd As String
Global Operacion As String
Global CodOper As String
Global MetReg As String
Global NombreProducto As String
Global Rut As String
Global Oficina As String
Public Sub Ilumina_Row(Grd_Sel As Control)
    If Grd_Sel.Row >= 0 Then
       Grd_Sel.SelStartCol = 0
       Grd_Sel.SelStartRow = Grd_Sel.Row
       Grd_Sel.SelEndCol = Grd_Sel.Cols - 1
       Grd_Sel.SelEndRow = Grd_Sel.Row
       Grd_Sel.HighLight = True
       Grd_Sel.Col = Grd_Sel.Cols - 1
    End If
End Sub
Public Function PoneValorINI(Grupo As String, Titulo As String, palabra As String, NombreArchivo As String) As Boolean
Dim RET As Integer
  PoneValorINI = False
  RET = WritePrivateProfileString(ByVal Grupo, ByVal Titulo, ByVal palabra, ByVal NombreArchivo)
  If RET = 1 Then
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

Public Sub ObtenerDatosDeCorte(ByVal Producto As String, IniLlave As Integer, LarLlave As Integer, IniData As Integer, LarData As Integer, NomProd As String, VistaDestino As String)
Dim PunteroTabPrd As Recordset
  Set PunteroTabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  PunteroTabPrd.Index = "Tab_Prd"
  PunteroTabPrd.Seek "=", Producto
  
  NomProd = PunteroTabPrd("Prd_Descrip")
  IniData = Interpreta(PunteroTabPrd("Prd_Pos_DataDst"), SIMBOLOPOSICION)
  IniLlave = Interpreta(PunteroTabPrd("Prd_Pos_LlaveDst"), SIMBOLOPOSICION)
  LarData = Interpreta(PunteroTabPrd("Prd_Pos_DataDst"), SIMBOLOLARGO)
  LarLlave = Interpreta(PunteroTabPrd("Prd_Pos_LlaveDst"), SIMBOLOLARGO)
  VistaDestino = PunteroTabPrd("Prd_Vista_Dst")
  
  PunteroTabPrd.Close
End Sub
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









Public Function Valida_3030(DataCTD As String) As Boolean

Dim Dat As String

Origen = IIf(Mid$(DataCTD, 217, 1) = "1", "BAE", "BCH")
TipoProd = "0" 'Cta Cte

'** VALIDA SEGUN TABLA
Call ValidarSegunTablas(0, "METRO", DataCTD)

'** VALIDA DATOS
Call Valida_CodCierre(Mid$(DataCTD, 158, 1), 158, "Causal de Cierre")
Call Valida_Rut(Mid$(DataCTD, 163, 9), 161, "Rut Cliente")
Call Valida_Cuenta(Origen, Mid$(DataCTD, 226, 12), 224, "N° Línea de Crédito")
Call Valida_Cuenta(Origen, Mid$(DataCTD, 266, 12), 264, "N° Cuenta Cte. Relacionada")

'Call Valida_Ejec(Mid$(186, 3), 186, "Código Ejecutivo")

Valida_3030 = True


End Function

Public Sub ValidarSegunTablas(Prod As Integer, Area As String, DatCTA As String)

Dim PunteroVld As Recordset
Dim PunteroReg As Recordset
Dim PunteroTbl As Recordset
Dim Tabla As String
Dim i As Integer
Dim ValorOri As String
Dim Encontrado As Boolean
Dim Descripcion1 As String
Dim Descripcion As String

Tabla = IIf(Area = "METRO", "tbl_ValMET", "tbl_ValREG")
     
Set PunteroVld = BaseVLDCTD.OpenRecordset("Select CTD_Orig,CTD_Categ,CTD_PosVista From " + Tabla + " Where CTD_Prod=" + Str(Prod) + " And CTD_Oper=" + CodOper + " Group by  CTD_Orig,CTD_Categ,CTD_PosVista  Order by CTD_PosVista Asc", dbOpenDynaset)
' PunteroVld:  campo categoría distintos

If PunteroVld.RecordCount > 0 Then
   PunteroVld.MoveLast
   PunteroVld.MoveFirst
   
   For i = 1 To PunteroVld.RecordCount
     
     Set PunteroReg = BaseVLDCTD.OpenRecordset("Select * From " + Tabla + " Where CTD_Prod=" + Str(Prod) + " And CTD_Oper=" + CodOper + _
                                           " And CTD_Orig=" + Str(PunteroVld!Ctd_Orig) + " And CTD_Categ='" + PunteroVld!CTD_Categ + "'", dbOpenDynaset)
     If PunteroReg.RecordCount > 0 Then
        Descripcion1 = Trim$(PunteroReg!Ctd_Descrip)
     End If
     PunteroReg.Close
     
     Set PunteroReg = BaseVLDCTD.OpenRecordset("Select * From " + Tabla + " Where CTD_Prod=" + Str(Prod) + " And CTD_Oper=" + CodOper + _
                                           " And CTD_Orig=" + Str(PunteroVld!Ctd_Orig) + " And CTD_Categ='" + PunteroVld!CTD_Categ + "'", dbOpenDynaset)
     
    ' PunteroReg: todos los que coinciden con la categoria
     
     PunteroReg.MoveLast   '(para saber el total por categoría)
     PunteroReg.MoveFirst
     If PunteroReg.RecordCount > 0 Then
       
       If PunteroReg.RecordCount = 1 Then
         '* Discriminar dato monto, o Fecha
         Select Case Mid$(PunteroReg!Ctd_TipoDat, 1, 1)
            Case "N"
               ValorOri = Mid$(DatCTA, PunteroReg!CTD_PosVista, PunteroReg!CTD_LargoVta)
               
               If Not IsNumeric(ValorOri) Then
                  If Mid$(PunteroReg!Ctd_TipoDat, 6, 1) = "I" Then
                    If Not Valida_y_Decodifica(ValorOri) Then
                      Call IngresaError(PunteroVld!CTD_PosVista, ValorOri, Descripcion1, 0)
                    End If
                  ElseIf Mid$(PunteroReg!Ctd_TipoDat, 6, 1) = "P" Then
                    Call Valida_Num(Mid$(PunteroReg!Ctd_TipoDat, 2, 5), ValorOri, PunteroReg!Ctd_Descrip, PunteroReg!Ctd_RangoMax, PunteroReg!Ctd_RangoMin)
                  Else
                    Call IngresaError(PunteroVld!CTD_PosVista, ValorOri, Descripcion1, 0)
                  End If
               
               ElseIf Val(PunteroReg!Ctd_RangoMax) > 0 Or Val(PunteroReg!Ctd_RangoMin) > 0 Then
                 Call Valida_Num(Mid$(PunteroReg!Ctd_TipoDat, 2, 5), ValorOri, PunteroReg!Ctd_Descrip, PunteroReg!Ctd_RangoMax, PunteroReg!Ctd_RangoMin)
               End If
                               
            Case "D"
               Call Valida_Fecha(Mid$(PunteroReg!Ctd_TipoDat, 7, 8), DatCTA, PunteroReg!CTD_PosVista, PunteroReg!Ctd_Descrip)
               
         End Select
         
       Else
                    
         If InStr(Descripcion1, "-") > 0 Then
           Descripcion1 = Mid$(Descripcion1, 1, InStr(Descripcion1, "-") - 1)
         End If
         
         ValorOri = Mid$(DatCTA, PunteroReg!CTD_PosVista, PunteroReg!CTD_LargoVta)
         
         Encontrado = False
         PunteroReg.MoveFirst
         For j = 1 To PunteroReg.RecordCount
            If PunteroReg!Ctd_Argum = Format(ValorOri, "000000") Then
               Encontrado = True
               Exit For
            End If
            PunteroReg.MoveNext
         Next j
          
         If Not Encontrado Then  '* No encuentra registro
           Call IngresaError(PunteroVld!CTD_PosVista, ValorOri, Descripcion1, 1)
         End If
         
       End If
     End If
     PunteroVld.MoveNext
   Next i
End If
PunteroVld.Close

End Sub



Public Function Valida_6391(DataCTD As String) As Boolean
Dim Dat As String

Origen = IIf(Mid$(DataCTD, 359, 1) = "1", "BAE", "BCH")
TipoProd = "0" 'Mid$(DataCTD, 263, 1)

'** VALIDA SEGUN TABLA
Call ValidarSegunTablas(0, "REGION", DataCTD)

'** VALIDA DATOS
Call Valida_Rut(Mid$(DataCTD, 16, 10), 16, "Rut Cliente")
Call Valida_EjeOrg(Mid$(DataCTD, 31, 6), 31, "Ejecutivo Control")
Call Valida_CodCierre(Mid$(DataCTD, 146, 1), 146, "Causal de Cierre")
Call Valida_Cuenta(Origen, Mid$(DataCTD, 367, 12), 367, "N° Línea de Crédito")
Call Valida_Cuenta(Origen, Mid$(DataCTD, 393, 12), 393, "N° Cuenta Relacionada")
Call Valida_Cuenta(Origen, Mid$(DataCTD, 424, 12), 424, "N° Línea Crédito Ex-BAE")

'Call Valida_Ejec(Mid$(DataCTD,28, 3), 28, "Ejecutivo Cuenta")
'Call Valida_Fecha("ddmmyyyy", Mid$(DataCTD, 65, 8), 79, "Fecha Apertura")

Valida_6391 = True

End Function

Function Strtran(ByVal st As String, ByVal cambia As String, ByVal por As String) As String
    Dim pos As Integer
    Dim lenCam As Integer
 
    lenCam = Len(cambia)
    pos = InStr(st, cambia)
    While pos > 0
        st = Mid$(st, 1, pos - 1) + por + Mid$(st, pos + lenCam)
        pos = InStr(pos + Len(por), st, cambia)
    Wend
    Strtran = st
 
End Function
 
Public Sub IngresaError(IngIni As String, IngDat As String, IngGlosa As String, IngTipo As Integer)

Dim TotalErr As Integer

Frm_ERVL.GridErr.Col = 0
Frm_ERVL.GridErr.Row = 1
If Frm_ERVL.GridErr.Rows = 2 And Frm_ERVL.GridErr.Text = "" Then
  Frm_ERVL.GridErr.Row = 1
Else
  Frm_ERVL.GridErr.Rows = Frm_ERVL.GridErr.Rows + 1
  Frm_ERVL.GridErr.Row = Frm_ERVL.GridErr.Rows - 1
End If
Frm_ERVL.GridErr.Col = 0
Frm_ERVL.GridErr.Text = IngGlosa
Frm_ERVL.GridErr.Col = 1
Frm_ERVL.GridErr.Text = IngDat
Frm_ERVL.GridErr.Col = 2
Frm_ERVL.GridErr.Text = IngIni
Frm_ERVL.GridErr.Col = 3
Frm_ERVL.GridErr.Text = IngTipo

End Sub

Public Function Valida_3031(DataCTD As String) As Boolean

Dim Dat As String

Origen = IIf(Mid$(DataCTD, 320, 1) = "1", "BAE", "BCH")
TipoProd = "1" 'Lin Cre

'** VALIDA SEGUN TABLA
Call ValidarSegunTablas(1, "METRO", DataCTD)

'** VALIDA DATOS
Call Valida_EjeOrg(Mid$(DataCTD, 112, 6), 112, "Código Ejecutivo Control")
Call Valida_CodCierre(Mid$(DataCTD, 173, 1), 173, "Código de Cierre")
Call Valida_Rut(Mid$(DataCTD, 182, 10), 182, "Rut Cliente")
Call Valida_Cuenta(Origen, Mid$(DataCTD, 330, 12), 330, "N° Cta. Cte. Asociada")

'Call Valida_Ejec(Mid$(DataCTD, 118, 3), 118, "Código Ejecutivo Cuenta")

Valida_3031 = True


End Function

