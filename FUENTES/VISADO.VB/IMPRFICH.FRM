VERSION 4.00
Begin VB.Form ImprimeFichaCredito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión Excepciones Visado"
   ClientHeight    =   1050
   ClientLeft      =   3120
   ClientTop       =   3705
   ClientWidth     =   3660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   1455
   Icon            =   "IMPRFICH.frx":0000
   Left            =   3060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   Top             =   3360
   Width           =   3780
   Begin Threed.SSFrame SSFrame1 
      Height          =   1005
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   3570
      _version        =   65536
      _extentx        =   6297
      _extenty        =   1773
      _stockprops     =   14
      forecolor       =   8388608
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
      Begin Threed.SSCommand Cancelar 
         Height          =   645
         Left            =   1845
         TabIndex        =   0
         Top             =   225
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Cancelar"
         forecolor       =   192
         font3d          =   1
         picture         =   "IMPRFICH.frx":030A
      End
      Begin Threed.SSCommand Imprimir 
         Height          =   645
         Left            =   990
         TabIndex        =   1
         Top             =   225
         Width           =   780
         _version        =   65536
         _extentx        =   1376
         _extenty        =   1138
         _stockprops     =   78
         caption         =   "&Imprimir"
         forecolor       =   8388608
         font3d          =   1
         autosize        =   1
         picture         =   "IMPRFICH.frx":0624
      End
   End
End
Attribute VB_Name = "ImprimeFichaCredito"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Nodo As String
Dim PunteroVsdHst As Snapshot
Private Function AgregarCeros(Numero As String, ByVal CantidadCeros As Integer) As String
  Numero = Trim$(Numero)
  For I% = 1 To (CantidadCeros - Len(Numero))
    Numero = "0" + Numero
  Next I%
  AgregarCeros = Numero
End Function
Private Function AgregarCaracter(Caracter As String, ByVal Cantidad As Integer) As String
Dim Cadena As String
  Cadena = ""
  For I% = 1 To Cantidad
    Cadena = Cadena + Caracter
  Next I%
  AgregarCaracter = Cadena
End Function
Private Function TransformaRut(Rut As String) As String
Dim ParteCentral As String
Dim ParteDerecha As String
Dim ParteIzquierda As String
Dim DigitoVerificador As String * 1
  Rut = DesarmaNumero(Rut, 10)
  If Rut <> "" Then
    ParteCentral = Mid(Rut, 4, 3)
    ParteDerecha = Mid(Rut, 7, 3)
    ParteIzquierda = Mid(Rut, 1, 3)
    DigitoVerificador = Right(Rut, 1)
    ParteIzquierda = ParteIzquierda + "." + ParteCentral + "." + ParteDerecha + "-" + DigitoVerificador
    While Val(Left(ParteIzquierda, 1)) < 1 Or Val(Left(ParteIzquierda, 1)) > 9
        ParteIzquierda = Mid(ParteIzquierda, 2)
    Wend
    While Len(ParteIzquierda) < 12
        ParteIzquierda = "0" + ParteIzquierda
    Wend
    TransformaRut = ParteIzquierda
  Else
    TransformaRut = " "
  End If
End Function
Private Function FormateaFecha(Fecha As String) As String
  If Val(Fecha) = 0 Then
    FormateaFecha = "  /  /    "
  Else
    FormateaFecha = Mid(Fecha, 7, 2) + "/" + Mid(Fecha, 5, 2) + "/" + Mid(Fecha, 1, 4)
  End If
End Function
Private Function DesarmaNumero(Numero As String, ByVal Largo As Integer) As String
Dim I As Integer
Dim NumeroTemporal As String
  NumeroTemporal = ""
  For I = 1 To Len(Numero)
    If InStr(".-", Mid$(Numero, I, 1)) = 0 Then
        NumeroTemporal = NumeroTemporal + Mid(Numero, I, 1)
    End If
  Next I
  For I = 1 To Largo - Len(NumeroTemporal)
    NumeroTemporal = "0" + NumeroTemporal
  Next I
  DesarmaNumero = NumeroTemporal
End Function
Private Sub Cancelar_Click()
  Unload Me
End Sub
Private Sub Cancelar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub
Private Sub Form_Activate()
  Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
  Screen.MousePointer = 11
  Principal.Enabled = False
  Call CentrarVentana(Me)
  
  '***
  Set BaseHistory = OpenDatabase(DIRECTORIO_01 + "Hist_vsd.mdb")
  Llena_Historico
  '***
  
  Set PunteroVsdHst = BaseHistory.CreateSnapshot("SELECT * FROM tbl_Hst WHERE Hst_Fech='" + _
                                                                                      Format(Date, "yyyymmdd") + "';")
  If Environ$(AMBIENTEDESARROLLO) = "DESA" Then
    Nodo = "U005"
  Else
    Nodo = "U025"
  End If
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  PunteroVsdHst.Close
  
  BaseHistory.Close
  
  Principal.Enabled = True
End Sub
Private Sub Imprimir_Click()
Dim Ret As String
Dim Pos As Integer
Dim Data As String
Dim Salto As Integer
Dim Linea As String
Dim Data2 As String
Dim Linea2 As String
Dim Linea1 As String
Dim Exc_To As String
Dim Numero As String
Dim Termino As Boolean
Dim LlaveSig As String
Dim Acc_Cre As String
Dim NumReg As Integer
Dim Rut_Cliente As String
Dim Nro_Pagare As String
Dim Ofi_Desemb As String
Dim NombreArchivo As String

Dim Pas1 As Boolean
Dim Pas2 As Boolean
Dim Pas4 As Boolean
Dim Pas6 As Boolean
Dim Pas8 As Boolean
Dim Pas9 As Boolean
Dim Pas13 As Boolean
Dim Pas15 As Boolean

Dim RutAutor10 As String
Dim FechRegul10 As String
Dim FechaCom10 As String

Dim RutAutor20 As String
Dim FechRegul20 As String
Dim FechaCom20 As String

Dim RutAutor40 As String
Dim FechRegul40 As String
Dim FechaCom40 As String

Dim RutAutor60 As String
Dim FechRegul60 As String
Dim FechaCom60 As String

Dim RutAutor80 As String
Dim FechRegul80 As String
Dim FechaCom80 As String

Dim RutAutor90 As String
Dim FechRegul90 As String
Dim FechaCom90 As String

Dim RutAutor13 As String
Dim FechRegul13 As String
Dim FechaCom13 As String

Dim RutAutor15 As String
Dim FechRegul15 As String
Dim FechaCom15 As String

On Error GoTo ErrorImpresion
  Screen.MousePointer = 11
  If PunteroVsdHst.RecordCount > 0 Then
    PunteroVsdHst.MoveFirst
    Open DIRECTORIO_01 + ARCHIVO_IMD01 For Output As #1
    While Not PunteroVsdHst.EOF
      If PunteroVsdHst("Hst_Tipo") = "CRE" Then
        Pas1 = False
        Pas2 = False
        Pas4 = False
        Pas6 = False
        Pas8 = False
        Pas9 = False
        Pas13 = False
        Pas15 = False
        Termino = False
        Exc_To = AgregarCeros(Mid(PunteroVsdHst("Hst_Data"), 27, 3), 5)
        Acc_Cre = Mid(PunteroVsdHst("Hst_Data"), 26, 1)
        Rut_Cliente = Mid(PunteroVsdHst("Hst_Data"), 10, 10)
        Nro_Pagare = AgregarCeros(Mid(PunteroVsdHst("Hst_Data"), 20, 6), 9)
        Ofi_Desemb = Mid(PunteroVsdHst("Hst_Data"), 1, 3)
        Data = RUTDEFECTO + "1828" + "L" + Ofi_Desemb + Rut_Cliente + Nro_Pagare + Acc_Cre + Exc_To + "000316"
        Retorno$ = BajadaPRN(Nodo, "SD20", Data, Ret)
        If Mid(Data, 1, 2) = "00" Or Mid(Data, 1, 2) = "96" Then
          Linea = Space(133)
          Mid(Linea, 1) = AgregarCaracter(Chr(196), 133)
          Mid(Linea, 1) = Chr(218)
          Mid(Linea, 70) = Chr(194)
          Mid(Linea, 105) = Chr(194)
          Mid(Linea, 133) = Chr(191)
          Print #1, SIMBOLOLETRACHICA + Linea
          
          Linea = Space(133)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 70) = Chr(179)
          Mid(Linea, 72) = "Fecha del Pagar" + Chr(130) + ": " + Format(Mid(PunteroVsdHst("Hst_Data"), 225, 8), "00/00/0000")
          Mid(Linea, 105) = Chr(179)
          Mid(Linea, 107) = "N" + Chr(167) + " del Pagar" + Chr(130) + ": " + Nro_Pagare
          Mid(Linea, 133) = Chr(179)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Space(133)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 70) = AgregarCaracter(Chr(196), 70)
          Mid(Linea, 70) = Chr(195)
          Mid(Linea, 105) = Chr(197)
          Mid(Linea, 133) = Chr(180)
          Print #1, SIMBOLOLETRACHICA + Linea
          
          Linea = Space(134)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 3) = "DECLARACION DEL ENCARGADO DEL VISADO"
          Mid(Linea, 70) = Chr(179)
          Mid(Linea, 72) = "N" + Chr(167) + " del T.O.: " + Exc_To
          Mid(Linea, 105) = Chr(179)
          Numero = Mid(PunteroVsdHst("Hst_Data"), 61, 16)
          Numero = Mid$(Numero, 1, 13) + "," + Mid$(Numero, 14) + "0"
          While Mid$(Numero, 1, 1) = "0"
            Numero = Mid$(Numero, 2)
          Wend
          Numero = ColocarSeparadorDeMiles(Numero)
          Mid(Linea, 107) = "Monto: " + Mid(Numero, 1, Len(Numero) - 5)
          Mid(Linea, 133) = Chr(179)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Space(133)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 70) = AgregarCaracter(Chr(196), 70)
          Mid(Linea, 70) = Chr(195)
          Mid(Linea, 105) = Chr(197)
          Mid(Linea, 133) = Chr(180)
          Print #1, SIMBOLOLETRACHICA + Linea
          
          Linea = Space(133)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 70) = Chr(179)
          Mid(Linea, 72) = "Acci" + Chr(162) + "n (**): " + Acc_Cre
          Mid(Linea, 105) = Chr(179)
          Mid(Linea, 107) = "Moneda: " + Mid(PunteroVsdHst("Hst_Data"), 39, 3)
          Mid(Linea, 133) = Chr(179)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Mid(Linea, 1) = AgregarCaracter(Chr(196), 133)
          Mid(Linea, 1) = Chr(192)
          Mid(Linea, 70) = Chr(193)
          Mid(Linea, 105) = Chr(193)
          Mid(Linea, 133) = Chr(217)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Mid(Linea, 1) = AgregarCaracter(Chr(196), 133)
          Mid(Linea, 1) = Chr(218)
          Mid(Linea, 6) = Chr(194)
          Mid(Linea, 67) = Chr(194)
          Mid(Linea, 77) = Chr(194)
          Mid(Linea, 90) = Chr(194)
          Mid(Linea, 118) = Chr(194)
          Mid(Linea, 133) = Chr(191)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Space(133)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 2) = "Cod."
          Mid(Linea, 6) = Chr(179)
          Mid(Linea, 8) = "Normas a Cumplir :"
          Mid(Linea, 67) = Chr(179)
          Mid(Linea, 68) = "Excepci" + Chr(162) + "n"
          Mid(Linea, 77) = Chr(179)
          Mid(Linea, 81) = "Fecha"
          Mid(Linea, 90) = Chr(179)
          Mid(Linea, 97) = "Nombre y Firma de "
          Mid(Linea, 118) = Chr(179)
          Mid(Linea, 122) = "Fecha en"
          Mid(Linea, 133) = Chr(179)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Space(133)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 6) = Chr(179)
          Mid(Linea, 8) = "(Entre par" + Chr(130) + "ntesis se referencian las normas del Anexo 2"
          Mid(Linea, 67) = Chr(179)
          Mid(Linea, 69) = "Si o No"
          Mid(Linea, 77) = Chr(179)
          Mid(Linea, 78) = "Comprometida"
          Mid(Linea, 90) = Chr(179)
          Mid(Linea, 98) = "Qui" + Chr(130) + "n Autoriza"
          Mid(Linea, 118) = Chr(179)
          Mid(Linea, 123) = "que se"
          Mid(Linea, 133) = Chr(179)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Space(133)
          Mid(Linea, 1) = Chr(179)
          Mid(Linea, 6) = Chr(179)
          Mid(Linea, 8) = "de la circular indicada)"
          Mid(Linea, 67) = Chr(179)
          Mid(Linea, 77) = Chr(179)
          Mid(Linea, 90) = Chr(179)
          Mid(Linea, 118) = Chr(179)
          Mid(Linea, 121) = "Regulariz" + Chr(162)
          Mid(Linea, 133) = Chr(179)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Mid(Linea, 1) = AgregarCaracter(Chr(196), 133)
          Mid(Linea, 1) = Chr(195)
          Mid(Linea, 6) = Chr(197)
          Mid(Linea, 67) = Chr(197)
          Mid(Linea, 77) = Chr(197)
          Mid(Linea, 90) = Chr(197)
          Mid(Linea, 118) = Chr(197)
          Mid(Linea, 133) = Chr(180)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Salto = 0
          NumReg = 1
          LlaveSig = Ofi_Desemb + Rut_Cliente + Nro_Pagare + Acc_Cre + Exc_To + "000"
          Linea1 = Linea
          While Not Termino
            Data2 = RUTDEFECTO + "1828" + "V" + "0000" + LlaveSig
            Retorno$ = BajadaPRN(Nodo, "SD20", Data2, Ret)
            While Left(LlaveSig, 28) = Mid(Data2, 12 + Salto, 28)
              If NumReg > 1 Then
                If Mid(Data2, 12 + Salto + 34, 1) = "2" Then
                  Select Case Mid(Data2, 12 + Salto + 28, 3)
                    Case "010"
                      Pas1 = True
                      RutAutor10 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul10 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom10 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                    Case "020"
                      Pas2 = True
                      RutAutor20 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul20 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom20 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                    Case "041", "042"
                      Pas4 = True
                      RutAutor40 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul40 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom40 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                    Case "060"
                      Pas6 = True
                      RutAutor60 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul60 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom60 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                    Case "080"
                      Pas8 = True
                      RutAutor80 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul80 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom80 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                    Case "090"
                      Pas9 = True
                      RutAutor90 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul90 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom90 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                    Case "130"
                      Pas13 = True
                      RutAutor13 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul13 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom13 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                    Case "150"
                      Pas15 = True
                      RutAutor15 = TransformaRut(Mid(Data2, 12 + Salto + 51, 10))
                      FechRegul15 = FormateaFecha(Mid(Data2, 12 + Salto + 61, 8))
                      FechaCom15 = FormateaFecha(Mid(Data2, 12 + Salto + 35, 8))
                  End Select
                End If
              End If
              Salto = Salto + 404
              NumReg = NumReg + 1
            Wend
            Pos = (Val(Mid(Data2, 4, 4)) - 1) * 404
            If Mid(Data2, Pos + 12, 28) <> Left(LlaveSig, 28) Then
              Termino = True
            Else
              LlaveSig = Mid(Data2, Pos + 12, 31)
              LlaveSig = Left(LlaveSig, 30) + Chr$(Asc(Right(LlaveSig, 1)) + 1)
              Salto = 0
            End If
          Wend
        
          Call PonerTituloDatos(Linea, "01", "Prohibici" + Chr(162) + "n y Margenes ", "(A y B)", 59)
          Call LlenarConDatos(Linea, Pas1, FechaCom10, RutAutor10, FechRegul10)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea1
        
          Call PonerTituloDatos(Linea, "02", "Vigencia de Sociedad", "(C)", 63)
          Call LlenarConDatos(Linea, Pas2, FechaCom20, RutAutor20, FechRegul20)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea1
        
          Call PonerTituloDatos(Linea, "04", "Pagar" + Chr(130) + " Bien Extendido y Consistente (*)", "(E)", 63)
          Call LlenarConDatos(Linea, Pas4, FechaCom40, RutAutor40, FechRegul40)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea1
         
          Call PonerTituloDatos(Linea, "06", "M.A.C. Solic.con Atribuciones y Facultades Suficientes", "(G)", 63)
          Call LlenarConDatos(Linea, Pas6, FechaCom60, RutAutor60, FechRegul60)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea1
        
          Call PonerTituloDatos(Linea, "08", "L" + Chr(161) + "nea de Credito Vigente y Margen Disponible", "(H)", 63)
          Call LlenarConDatos(Linea, Pas8, FechaCom80, RutAutor80, FechRegul80)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea1
        
          Call PonerTituloDatos(Linea, "09", "Cliente Creado y con Ejecutivo Control en C.D.R.", "(I)", 63)
          Call LlenarConDatos(Linea, Pas9, FechaCom90, RutAutor90, FechRegul90)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea1
        
          Call PonerTituloDatos(Linea, "13", "Garant" + Chr(161) + "as Cumplen con las Normas Exigidas", "(L)", 63)
          Call LlenarConDatos(Linea, Pas13, FechaCom13, RutAutor13, FechRegul13)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea1
        
          Call PonerTituloDatos(Linea, "15", "Certificaci" + Chr(162) + "n Notarial de Pagar" + Chr(130) + " en Orden", "(D)", 63)
          Call LlenarConDatos(Linea, Pas15, FechaCom15, RutAutor15, FechRegul15)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Mid(Linea, 1) = AgregarCaracter(Chr(196), 133)
          Mid(Linea, 1) = Chr(195)
          Mid(Linea, 6) = Chr(193)
          Mid(Linea, 67) = Chr(193)
          Mid(Linea, 77) = Chr(193)
          Mid(Linea, 90) = Chr(193)
          Mid(Linea, 118) = Chr(193)
          Mid(Linea, 133) = Chr(180)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea1 = Space(133)
          Mid(Linea1, 1) = Chr(179)
          Mid(Linea1, 133) = Chr(179)
        
          Linea = Linea1
          Mid(Linea, 2 _
                                  ) = "(a)   Si se detectan excepciones a las normas y est" + Chr(160) + "n autorizadas en la Declaraci" + Chr(162) + "n del Ejecutivo, debe indicarse NO."
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea1
          Mid(Linea, 2 _
                                 ) = "(b)   Si se detectan excepciones a las normas y no est" + Chr(160) _
                                 + "n autorizadas en la Declaraci" + Chr(162) + "n del Ejecutivo, " + Chr(130) + "stas se deben autorizar en "
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea1
          Mid(Linea, 8) = "esta secci" + Chr(162) + "n."
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Mid(Linea, 1) = AgregarCaracter(Chr(196), 133)
          Mid(Linea, 1) = Chr(195)
          Mid(Linea, 133) = Chr(180)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea2 = Linea
          Linea = Linea1
          Mid(Linea, 2) = "Observaciones:"
          Mid(Linea, 18) = Mid(Data, 139, 50)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea1
          Mid(Linea, 18) = Mid(Data, 189, 50)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRACHICA + Linea2
          Print #1, SIMBOLOLETRACHICA + Linea1
          Print #1, SIMBOLOLETRACHICA + Linea1
          Print #1, SIMBOLOLETRACHICA + Linea1
          Print #1, SIMBOLOLETRACHICA + Linea1
        
          Linea = Linea1
          Mid(Linea, 2) = "C" + Chr(162) + "digo Oficina: " + Environ$(AMBIENTEOFICINA)
          Mid(Linea, 85) = AgregarCaracter(Chr(196), 45)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea1
          Mid(Linea, 2) = "Fecha: " + Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date))
          Mid(Linea, 88) = "Timbre y Firma del Encargado de Visado"
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea2
          Mid(Linea, 1) = Chr(192)
          Mid(Linea, 133) = Chr(217)
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Mid(Linea, 1) = Chr(218)
          Mid(Linea, 133) = Chr(191)
          Print #1, SIMBOLOLETRACHICA + Linea
          
          Linea = Linea1
          Mid(Linea, 2) = "(**)"
          Mid(Linea, 7) = "C" + Chr(162) + "digos de Acci" + Chr(162) + "n :"
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea1
          Mid(Linea, 7) = "F = Creaci" + Chr(162) + "n/Modificaci" + Chr(162) + "n de Cuadros de Pago"
          Mid(Linea, 70) = "N = Creaci" + Chr(162) + "n de un Pr" + Chr(130) + "stamo"
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea1
          Mid(Linea, 7) = "M = Modificaci" + Chr(162) + "n de Cartera"
          Mid(Linea, 70) = "R = Renovaci" + Chr(162) + "n de un Pr" + Chr(130) + "stamo"
          Print #1, SIMBOLOLETRACHICA + Linea
        
          Linea = Linea2
          Mid(Linea, 1) = Chr(192)
          Mid(Linea, 133) = Chr(217)
          Print #1, SIMBOLOLETRACHICA + Linea
          Print #1, SIMBOLOLETRANORMAL + Chr(12)
        Else
          MsgBox "No Hay Comunicación con el Servidor", 0 + 16, "Problemas de Comunicación"
          Close #1
          Kill DIRECTORIO_01 + ARCHIVO_IMD01
          Screen.MousePointer = 0
          On Error GoTo 0
          Exit Sub
        End If
      End If
      PunteroVsdHst.MoveNext
    Wend
    Close #1
    
    If Environ("OS") = "" Then ' Win 3.11
      NombreArchivo = DIRECTORIO_01 + ARCHIVO_PIF01
      If Dir(NombreArchivo) = UCase(Mid(NombreArchivo, Len(NombreArchivo) - 7)) Then
        Call EjecutarPrograma(NombreArchivo + " " + DIRECTORIO_01 + ARCHIVO_IMD01 + " " + DefinicionImpresora, 0)
      Else
        MsgBox "No se Puede Ejecutar Programa" + Chr(10) + NombreArchivo, 0 + 48, "Falta Programa Asociado"
      End If
    Else  ' Win NT, Win 2000
      NombreArchivo = DIRECTORIO_02 + "IMPRES.EXE"
      If Dir(NombreArchivo) = "IMPRES.EXE" Then
        Call EjecutarPrograma(NombreArchivo + " " + DIRECTORIO_01 + ARCHIVO_IMD01 + " " + DefinicionImpresora, 0)
      Else
        MsgBox "No se Puede Ejecutar Programa" + Chr(10) + NombreArchivo, 0 + 48, "Falta Programa Asociado"
      End If
    End If
    
    Kill DIRECTORIO_01 + ARCHIVO_IMD01
  Else
    MsgBox "No se Encontraron Datos a Imprimir", 0 + 16, "Faltan Datos"
  End If
  Screen.MousePointer = 0
On Error GoTo 0
  Exit Sub

ErrorImpresion:
  MsgBox "Grave Error con la Impresora", 0 + 16, "Falla de Impresión"
  Close #1
  Kill DIRECTORIO_01 + ARCHIVO_IMD01
On Error GoTo 0
  Screen.MousePointer = 0
  Exit Sub


End Sub
Private Sub Imprimir_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub


Private Sub LlenarConDatos(Datos As String, ByVal Estado As Boolean, ByVal FechaCompromiso As String, ByVal RutAutorizacion As String, ByVal FechaRegularizacion As String)
  If Not Estado Then
    Mid(Datos, 71) = "NO"
  Else
    Mid(Datos, 71) = "SI"
    Mid(Datos, 80) = FechaCompromiso
    Mid(Datos, 92) = RutAutorizacion
    Mid(Datos, 120) = FechaRegularizacion
  End If
End Sub

Private Sub PonerTituloDatos(Datos As String, ByVal Numero As String, ByVal Glosa As String, ByVal Puntos As String, ByVal PosicionPuntos As Integer)
  Datos = Space(133)
  Mid(Datos, 1) = Chr(179)
  Mid(Datos, 6) = Chr(179)
  Mid(Datos, 67) = Chr(179)
  Mid(Datos, 77) = Chr(179)
  Mid(Datos, 90) = Chr(179)
  Mid(Datos, 118) = Chr(179)
  Mid(Datos, 133) = Chr(179)
  Mid(Datos, 3) = Numero
  Mid(Datos, 8) = Glosa
  Mid(Datos, PosicionPuntos) = Puntos
End Sub

