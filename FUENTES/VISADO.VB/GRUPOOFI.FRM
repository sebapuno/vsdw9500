VERSION 4.00
Begin VB.Form GrupoOficina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configura Grupo Oficina"
   ClientHeight    =   5115
   ClientLeft      =   2775
   ClientTop       =   2130
   ClientWidth     =   5280
   ControlBox      =   0   'False
   Height          =   5520
   Left            =   2715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Top             =   1785
   Width           =   5400
   Begin Threed.SSFrame SSFrame1 
      Height          =   5055
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _version        =   65536
      _extentx        =   9128
      _extenty        =   8916
      _stockprops     =   14
      forecolor       =   32768
      BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      font3d          =   1
      shadowstyle     =   1
      Begin VB.ListBox OfiSelec 
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1500
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   4935
      End
      Begin VB.ListBox ListaOficinas 
         BeginProperty Font 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1500
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   4935
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4935
         _version        =   65536
         _extentx        =   8705
         _extenty        =   450
         _stockprops     =   15
         caption         =   "LISTADO DE OFICINAS"
         backcolor       =   12632256
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         font3d          =   2
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   4935
         _version        =   65536
         _extentx        =   8705
         _extenty        =   450
         _stockprops     =   15
         caption         =   "OFICINAS SELECCIONADAS"
         backcolor       =   12632256
         BeginProperty font {FB8F0823-0164-101B-84ED-08002B2EC713} 
            name            =   "MS Sans Serif"
            charset         =   1
            weight          =   400
            size            =   9.75
            underline       =   0   'False
            italic          =   0   'False
            strikethrough   =   0   'False
         EndProperty
         font3d          =   2
      End
      Begin Threed.SSCommand btn_Eliminar 
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   4280
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Eliminar"
         forecolor       =   128
         font3d          =   3
         picture         =   "GRUPOOFI.frx":0000
      End
      Begin Threed.SSCommand btn_Agregar 
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   4280
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Agregar"
         forecolor       =   12582912
         font3d          =   3
         picture         =   "GRUPOOFI.frx":031A
      End
      Begin Threed.SSCommand Aceptar 
         Height          =   615
         Left            =   3120
         TabIndex        =   2
         Top             =   4275
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Aceptar"
         forecolor       =   8421376
         font3d          =   3
         picture         =   "GRUPOOFI.frx":0634
      End
      Begin Threed.SSCommand Cancelar 
         Height          =   615
         Left            =   4080
         TabIndex        =   1
         Top             =   4275
         Width           =   855
         _version        =   65536
         _extentx        =   1508
         _extenty        =   1085
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   3
         picture         =   "GRUPOOFI.frx":094E
      End
   End
End
Attribute VB_Name = "GrupoOficina"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Aceptar_Click()

Dim BaseOficina    As Database
Dim PunteroOficina As Recordset
Dim BaseSgt As Database
Dim PunteroSgt As Recordset
  
Screen.MousePointer = 11
Set BaseOficina = OpenDatabase(DIRECTORIO_01 + ARCHIVO_MDB03)
Set BaseSgt = OpenDatabase("c:\data\admsgt\tabsgt.mdb")
BaseOficina.Execute "Delete * from GRP_OFI where grp_oficui='" + Environ("OFICINA") + "'"

If OfiSelec.ListCount > 0 Then
  Set PunteroOficina = BaseOficina.OpenRecordset("GRP_OFI", dbOpenTable)
  For I = 0 To OfiSelec.ListCount - 1
    Set PunteroSgt = BaseSgt.OpenRecordset("select * from tblcui where cui_cuicod=" + Mid(OfiSelec.List(I), 1, 3), dbOpenDynaset)
    If PunteroSgt.RecordCount > 0 Then
      If Val(Environ("OFICINA")) <> Val(PunteroSgt!cui_cuisuc) Then
        PunteroOficina.AddNew
        PunteroOficina!GRP_OFICOD = Mid$(OfiSelec.List(I), 1, 3)
        PunteroOficina!GRP_OFICUI = Environ("OFICINA")
        PunteroOficina.Update
      End If
    Else
      MsgBox "Oficina " + Trim(OfiSelec.List(I)) + " no se encuentra en tabla SGT." + _
               Chr(10) + "No se puede asociar la oficina.", 16, "Oficina No Encontrada"
    End If
  Next I
  PunteroSgt.Close
  PunteroOficina.Close
End If
BaseSgt.Close
BaseOficina.Close

MsgBox "Los datos del Grupo de Oficinas han sido actualizados.", 64
Unload Me

Screen.MousePointer = 0

End Sub


Private Sub btn_Agregar_Click()

Dim Encontrado As Boolean

If ListaOficinas.ListIndex < 0 Then
  Exit Sub
End If
Screen.MousePointer = 11

Encontrado = False
If OfiSelec.ListCount > 0 Then
  For I = 0 To OfiSelec.ListCount - 1
     If ListaOficinas.Text = OfiSelec.List(I) Then
        MsgBox "La Oficina ya está seleccionada.", 64
        Encontrado = True
        Exit For
     End If
  Next I
End If

If Not Encontrado Then
  If Not CoincideRegion(Mid$(ListaOficinas.Text, 1, 3)) Then
    valen = MsgBox("La oficina seleccionada (" + Mid$(ListaOficinas.Text, 1, 3) + ") pertenece a una región distinta a la correspondiente al PC." + _
                     Chr(10) + "Esto puede generar problemas al realizar consultas de visado." + Chr(10) + "¿Desea agregarla de todas formas?", 52, "Advertencia")
  Else
    valen = vbYes ' TIENE MISMA REGION
  End If
  
  If valen = vbYes Then
    OfiSelec.AddItem ListaOficinas.Text
    If OfiSelec.NewIndex > 6 Then
      OfiSelec.TopIndex = OfiSelec.NewIndex - 5
    End If
    OfiSelec.Selected(OfiSelec.NewIndex) = True
  End If
End If

Screen.MousePointer = 0

End Sub


Private Sub btn_Eliminar_Click()

Dim BaseSgt As Database
Dim PunteroSgt As Recordset
Dim NumeroLista As Integer

Set BaseSgt = OpenDatabase("c:\data\admsgt\tabsgt.mdb")

If OfiSelec.ListIndex > -1 Then
  Set PunteroSgt = BaseSgt.OpenRecordset("select * from tblcui where cui_cuicod=" + Mid(OfiSelec.Text, 1, 3), dbOpenDynaset)
  If PunteroSgt.RecordCount > 0 Then
    If Val(Environ("OFICINA")) <> Val(PunteroSgt!cui_cuisuc) Then
      NumeroLista = OfiSelec.ListIndex
      OfiSelec.RemoveItem (NumeroLista)
      If NumeroLista > 0 Then
        OfiSelec.Selected(NumeroLista - 1) = True
      Else
        OfiSelec.Selected(0) = True
      End If
    Else
      MsgBox "La oficina está relacionada con la correspondiente al PC.", 48, "No se puede eliminar la oficina"
    End If
  Else
    NumeroLista = OfiSelec.ListIndex
    OfiSelec.RemoveItem (NumeroLista)
    If NumeroLista > 0 Then
      OfiSelec.Selected(NumeroLista - 1) = True
    Else
      OfiSelec.Selected(0) = True
    End If
  End If
  PunteroSgt.Close
End If
 
BaseSgt.Close
 
End Sub

Private Sub Cancelar_Click()
 Unload Me
End Sub


Private Sub Form_Load()

Screen.MousePointer = 11
Principal.Enabled = False
Call CentrarVentana(Me)

CargarListaOficinas

CargarOficinasSeleccionadas

Screen.MousePointer = 0

End Sub



Public Sub CargarListaOficinas()

Dim BaseSucursales    As Database
Dim PunteroSucursales As Recordset
  
Set BaseSucursales = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
Set PunteroSucursales = BaseSucursales.OpenRecordset("Select Distinct SUC_SUCCOD,SUC_SUCNOM from TblSuc", dbOpenDynaset)
If PunteroSucursales.RecordCount > 0 Then
    PunteroSucursales.MoveLast
    PunteroSucursales.MoveFirst
    For I = 1 To PunteroSucursales.RecordCount
      ListaOficinas.AddItem Format(PunteroSucursales!SUC_SUCCOD, "000") + Space(4) + PunteroSucursales!Suc_SucNom
      PunteroSucursales.MoveNext
    Next I
    ListaOficinas.ListIndex = 0
End If
PunteroSucursales.Close
BaseSucursales.Close

  
End Sub

Private Sub Form_Unload(Cancel As Integer)

Principal.Enabled = True

End Sub


Private Sub ListaOficinas_DblClick()

btn_Agregar_Click

End Sub



Public Sub CargarOficinasSeleccionadas()

Dim BaseOficina As Database
Dim PunteroOficina As Recordset
Dim Encontro As Boolean

Set BaseOficina = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
Set PunteroOficina = BaseOficina.OpenRecordset("SELECT DISTINCT CUI_CUICOD FROM TBLCUI WHERE CUI_CUISUC=" + Environ("OFICINA") + " AND CUI_CUIISU='S'", dbOpenDynaset)
While Not PunteroOficina.EOF
  OfiSelec.AddItem GlosaOficina(Format(PunteroOficina!CUI_CUICOD, "000"))
  PunteroOficina.MoveNext
Wend
PunteroOficina.Close
BaseOficina.Close

Set BaseOficina = OpenDatabase(DIRECTORIO_01 + "OFICINA.MDB")
Set PunteroOficina = BaseOficina.OpenRecordset("Select * From GRP_OFI Where GRP_OFICUI='" + Environ("OFICINA") + "' ", dbOpenDynaset)
While Not PunteroOficina.EOF
  For I = 0 To OfiSelec.ListCount - 1
     If PunteroOficina!GRP_OFICOD = Mid$(OfiSelec.List(I), 1, 3) Then
       Encontro = True
       Exit For
     End If
  Next I
  If Not Encontro And PunteroOficina!GRP_OFICUI = Environ("OFICINA") Then
    OfiSelec.AddItem GlosaOficina(Format(PunteroOficina!GRP_OFICOD, "000"))
  End If
  PunteroOficina.MoveNext
Wend
PunteroOficina.Close
BaseOficina.Close

If OfiSelec.ListCount > 0 Then
  OfiSelec.Selected(0) = True
End If

End Sub





Private Sub ListaOficinas_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  btn_Agregar_Click
End If

End Sub

Private Sub OfiSelec_DblClick()

btn_Eliminar_Click

End Sub




Public Function CoincideRegion(CodOfiIng As String) As Boolean

Dim BaseSucursales    As Database
Dim PunteroSucursales As Recordset
     
CoincideRegion = False

Set BaseSucursales = OpenDatabase("C:\DATA\ADMSGT\TABSGT.MDB")
Set PunteroSucursales = BaseSucursales.OpenRecordset("Select * from Tblcui where cui_cuicod=" + Environ("OFICINA"), dbOpenDynaset)
If PunteroSucursales.RecordCount > 0 Then
  REGION_OFI = PunteroSucursales!cui_cuireg
  Set PunteroSucursales = BaseSucursales.OpenRecordset("Select * from Tblcui where cui_cuicod=" + CodOfiIng, dbOpenDynaset)
  If PunteroSucursales.RecordCount > 0 Then
    If REGION_OFI = PunteroSucursales!cui_cuireg Then
       CoincideRegion = True
    End If
  End If
End If

PunteroSucursales.Close
BaseSucursales.Close

End Function
