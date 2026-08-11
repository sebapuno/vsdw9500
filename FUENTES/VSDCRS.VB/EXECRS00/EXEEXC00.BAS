Attribute VB_Name = "ExcVisado"
Option Explicit
Dim Ap_VsdPrd As Recordset
Dim Ap_TabOper As Recordset
Dim Ap_TabPrd As Recordset

Global GetAis As String * 70

Public Function ExcepcionesAlVisado()
Dim Actividad As String * 1
Dim Nom_Arch As String
Dim Ap_Accion As String * 1
Dim ap_rut  As String
Dim ap_correlativo As String
Dim ap_producto As String

  Set Ap_VsdPrd = BaseVisado.OpenRecordset("Vsd_prd", dbOpenTable)
  ' FI2026 - era: Ap_VsdPrd.Index = Ap_VsdPrd.Indexes(0).Name
  ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
  ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
  ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
  Ap_VsdPrd.Index = BaseVisado.TableDefs("Vsd_prd").Indexes(0).Name
  
  Set Ap_TabOper = BaseVisado.OpenRecordset("Tab_Oper", dbOpenTable)
  ' FI2026 - era: Ap_TabOper.Index = Ap_TabOper.Indexes(0).Name
  ' Recordset.Indexes esta [restricted] en DAO 3.6, pero TableDef.Indexes NO.
  ' NO es un no-op: sin indice activo el .Seek de mas abajo falla en runtime
  ' con error 3019. dbOpenTable deja .Index vacio (verificado en la VM).
  Ap_TabOper.Index = BaseVisado.TableDefs("Tab_Oper").Indexes(0).Name
  
  Set Ap_TabPrd = BaseVisado.OpenRecordset("Tab_Prd", dbOpenTable)
  Ap_TabPrd.Index = "TAB_PRD"
  
  '*** Obtiene Parametros
  ap_rut = SacaValorINI("PARAMETROS", "PARAM1", "C:\Data\Visado\Visado.INI")
  ap_producto = SacaValorINI("PARAMETROS", "PARAM2", "C:\Data\Visado\Visado.INI")
  ap_correlativo = SacaValorINI("PARAMETROS", "PARAM3", "C:\Data\Visado\Visado.INI")
  Ap_Accion = SacaValorINI("PARAMETROS", "PARAM4", "C:\Data\Visado\Visado.INI")
  
  If Ap_Accion <> "0" Then
  
  End If
  
End Function
