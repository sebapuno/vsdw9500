VERSION 4.00
Begin VB.Form Frm_BajCTD 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualización de Parámetros CTD LCD "
   ClientHeight    =   900
   ClientLeft      =   3120
   ClientTop       =   3405
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   1305
   Left            =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   Top             =   3060
   Width           =   3810
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Espere un momento"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Index           =   2
      Left            =   975
      TabIndex        =   1
      Top             =   525
      Width           =   1815
   End
   Begin VB.Label lb_Baja 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3690
   End
End
Attribute VB_Name = "Frm_BajCTD"
Attribute VB_Creatable = False
Attribute VB_Exposed = False





Private Sub Form_Load()

On Error GoTo ErrorBajada

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Me.Show
Me.Refresh

Set BaseVisado = OpenDatabase("C:\DATA\VISADO\VISADO.MDB")
Set BaseEnvioSRM = OpenDatabase("C:\DATA\VISADO\ENVIOSRM.MDB")
Set BaseVldCta = OpenDatabase("C:\DATA\VISADO\VALCTDLC.MDB")

IdUsuario = ObtieneRutUsuarioLogueado()

BaseVldCta.Execute "Delete * From tbl_ValMET"
BaseVldCta.Execute "Delete * From tbl_ValREG"

Terminar = False

Bajar_Datos ("4086")
Bajar_Datos ("4087")

BaseVisado.Close
BaseEnvioSRM.Close
BaseVldCta.Close

End

ErrorBajada:

MsgBox "Se produjo un error durante la actualización de datos." + Chr(10) + Chr(10) + _
       Str(Err.Number) + Space(2) + Err.Description, 16, "Problema en Tabla de Validación"

BaseVisado.Close
BaseEnvioSRM.Close
BaseVldCta.Close

End

End Sub

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
          
          lb_Baja.Caption = Mid$(CadenaDatos, 15, 4) + Space(3) + Mid$(CadenaDatos, 19 + 6, 20)
          lb_Baja.Refresh
          
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

