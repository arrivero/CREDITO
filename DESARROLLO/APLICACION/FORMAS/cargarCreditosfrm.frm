VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form cargarCreditosfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevos Créditos desde Hand Held"
   ClientHeight    =   5310
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbCobrador 
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin FPSpread.vaSpread sprCreditos 
      Height          =   3885
      Left            =   30
      TabIndex        =   2
      Top             =   750
      Width           =   12435
      _Version        =   196608
      _ExtentX        =   21934
      _ExtentY        =   6853
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   10
      SpreadDesigner  =   "cargarCreditosfrm.frx":0000
   End
   Begin VB.CommandButton cmdregistra 
      Caption         =   "Registrar Créditos"
      Height          =   495
      Left            =   9630
      TabIndex        =   1
      Top             =   4740
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   11100
      TabIndex        =   0
      Top             =   4740
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cmdListaCreditos 
      Left            =   4500
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   10
   End
   Begin Crystal.CrystalReport crFactura 
      Left            =   10050
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtFechaRegistro 
      Height          =   345
      Left            =   7380
      TabIndex        =   5
      Top             =   240
      Width           =   1965
      _Version        =   65537
      _ExtentX        =   3466
      _ExtentY        =   609
      _StockProps     =   93
      Enabled         =   0   'False
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin VB.Label Label11 
      Caption         =   "Fecha:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6420
      TabIndex        =   6
      Top             =   300
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Cobrador:"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   270
      Width           =   1335
   End
End
Attribute VB_Name = "cargarCreditosfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_CREDITOS_SELECCIONADO = 1
Private Const COL_CREDITOS_CLIENTE = 2
Private Const COL_CREDITOS_NOMBRE_CLIENTE = 3
Private Const COL_CREDITOS_CREDITO = 4
Private Const COL_CREDITOS_INTERES = 5
Private Const COL_CREDITOS_FINANCIAMIENTO = 6
Private Const COL_CREDITOS_NO_PAGOS = 7
Private Const COL_CREDITOS_CANT_PAGAR = 8
Private Const COL_CREDITOS_TOTAL_PAGAR = 9
Private Const COL_CREDITOS_AUTORIZADO = 10

Private iVeces As Integer

Private Sub Form_Load()

    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        fnLlenaComboCollecion cbCobrador, oUsuario.cDatos, 0, ""
    End If
    Set oUsuario = Nothing

    'Obten los nuevos créditos desde la HH
    Call cargaCreditosHH
    
End Sub

Private Function cargaCreditosHH() As Boolean

    Dim cCreditos As Collection
    Dim bCreditos As Boolean
    
    cargaCreditosHH = False
    
    Set cCreditos = obtenCreditosDesdeArchivoHH(bCreditos)
    
    If bCreditos = True Then
        
        Call fnLlenaTablaCollection(sprCreditos, cCreditos)
        
        Call fnIdentificaCreditosNoAutorizados
        
        cargaCreditosHH = True
        
    End If
    
End Function

Private Function obtenCreditosDesdeArchivoHH(ByRef bCreditos As Boolean) As Collection
    
    On Error GoTo ErrArchivo
    
    Dim strArchivo As String
    Dim iPreciosArchivo As Integer
    
    bCreditos = False
    
    cmdListaCreditos.Filter = "CreditosNuevos(*.txt)|*.txt"
    cmdListaCreditos.FileName = "CreditosNuevos"
    cmdListaCreditos.DialogTitle = "Importar Créditos Nuevos"
    cmdListaCreditos.ShowOpen
    
    strArchivo = cmdListaCreditos.FileName
    
    'LLevar los códigos a la base de datos
    If abreArchivofn(strArchivo, iPreciosArchivo, PARA_LECTURA) Then

        Dim Registros As New Collection
        Dim Registro As Collection
        Dim oCampo As New Campo
        Dim strCampo As String
        Dim iCliente As Integer
        Dim strRegistro As String
        Dim fCantidad As Double
        
        Dim iPosicion As Integer
        Dim iPosicionComa As Integer
        
        Dim oCliente As New Cliente
        Dim oCredito As New credito
        
        Do While Not EOF(iPreciosArchivo)
            
            Set Registro = New Collection
            
            strRegistro = ""
            
            obtenRegistrofn iPreciosArchivo, strRegistro
            
            iPosicion = 1
            
            Registro.Add oCampo.CreaCampo(adInteger, , , 0)
            
            Do
                
                iPosicionComa = InStr(iPosicion, strRegistro, ",")
                
                If iPosicionComa > 0 Then
                
                    strCampo = Mid(strRegistro, iPosicion, iPosicionComa - iPosicion)
                    
                    Registro.Add oCampo.CreaCampo(adInteger, , , strCampo)
                    
                    If iPosicion = 1 Then
                        
                        iCliente = Val(strCampo)
                    
                        'Estos datos son para complemento de la operación
                        If oCliente.fnInformacion(iCliente, Date) = True Then
                            
                            Registro.Add oCampo.CreaCampo(adInteger, , , oCliente.mstrNombre + " " + oCliente.mstrApPaterno)    'Cliente
                            
                        End If
                    
                    End If
                    
                    iPosicion = iPosicionComa + 1
                    
                End If
                    
            Loop Until iPosicionComa = 0
            
            strCampo = Mid(strRegistro, iPosicion, Len(strRegistro) - (iPosicion - 1))
            Registro.Add oCampo.CreaCampo(adInteger, , , strCampo) 'Cantidad
                       
            fCantidad = Val(strCampo)
                       
            Registro.Add oCampo.CreaCampo(adInteger, , , 14) 'Interes
            Registro.Add oCampo.CreaCampo(adInteger, , , fCantidad * (14 / 100)) 'Financiamiento
            Registro.Add oCampo.CreaCampo(adInteger, , , 30) 'No. de Pagos
            Registro.Add oCampo.CreaCampo(adInteger, , , (fCantidad + (fCantidad * (14 / 100))) / 30) 'Cantidad a Pagar
            Registro.Add oCampo.CreaCampo(adInteger, , , fCantidad + (fCantidad * (14 / 100))) 'Cantidad Total
                       
            If True = oCredito.validaDisponibilidadDeCredito(iCliente, fCantidad) Then
                Registro.Add oCampo.CreaCampo(adInteger, , , 1) 'Disponibilidad de crédito
            Else
                Registro.Add oCampo.CreaCampo(adInteger, , , 0) 'No disponibilidad de crédito
            End If
                       
            Registros.Add Registro
            
            bCreditos = True
            
        Loop
        
        Set obtenCreditosDesdeArchivoHH = Registros
        
        cierraArchivofn iPreciosArchivo
        
        Set oCliente = Nothing
        Set oCredito = Nothing
        
    End If
    
ErrArchivo:
    Exit Function

End Function

Private Function fnIdentificaCreditosNoAutorizados()

    Dim lRow As Long
    
    sprCreditos.Col = COL_CREDITOS_AUTORIZADO
    
    For lRow = 1 To sprCreditos.DataRowCnt
    
        sprCreditos.Row = lRow
        If Val(sprCreditos.Text) = 0 Then
            
            ' Lock block of cells
            ' Specify the block of cells
            sprCreditos.Col = 1
            sprCreditos.Col2 = -1
            sprCreditos.Row = lRow
            sprCreditos.Row2 = lRow
            ' Lock cells
            sprCreditos.Lock = True
            
            sprCreditos.LockForeColor = RGB(255, 0, 0)
            
            sprCreditos.Lock = False
            
        Else
        
            ' Lock block of cells
            ' Specify the block of cells
            sprCreditos.Col = 1
            sprCreditos.Col2 = -1
            sprCreditos.Row = lRow
            sprCreditos.Row2 = lRow
            ' Lock cells
            sprCreditos.Lock = True
            
            sprCreditos.LockForeColor = RGB(0, 0, 0)
            
            sprCreditos.Lock = False
            
        End If
        
    Next lRow
    
End Function

Private Sub sprCreditos_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case vbKeyDelete
            sprCreditos.Col = sprCreditos.ActiveCol
            sprCreditos.Row = sprCreditos.ActiveRow
            sprCreditos.Action = ActionDeleteRow
        
        Case vbKeyF7
            sprCreditos.Col = sprCreditos.ActiveCol
            sprCreditos.Row = sprCreditos.ActiveRow
            sprCreditos.Action = ActionDeleteRow
    
    End Select
    
End Sub

Private Function fnValidaForma() As Boolean
    
    fnValidaForma = True
    
    If cbCobrador.Text = "" Then
        MsgBox "Seleccione un cobrador", vbInformation + vbOKOnly
        cbCobrador.SetFocus
        fnValidaForma = False
        Exit Function
    End If
    
    Dim lRow As Long
    Dim bCreditoSeleccionado As Boolean
    
    sprCreditos.Col = COL_CREDITOS_SELECCIONADO '1
    For lRow = 1 To sprCreditos.DataRowCnt
    
        sprCreditos.Row = lRow
        
        If sprCreditos.Text = 1 Then
            bCreditoSeleccionado = True
        End If
        
    Next lRow
    
    If bCreditoSeleccionado = False Then
        
        MsgBox "¡Para registrar creditos, seleccione almenos uno!", vbInformation + vbOKOnly
        sprCreditos.SetFocus
        fnValidaForma = False
        Exit Function
        
    End If
    
End Function

Private Sub cmdregistra_Click()

    If fnValidaForma = False Then
        Exit Sub
    End If
    
    accesofrm.Show vbModal

    'Si fue aceptada (si la propiedad bPermiteAcceso de la ventana 'accesofrm' es = true), hacer lo siguiente
    If accesofrm.bPermiteAcceso = True Then
        
        Dim oCredito As New credito
        
        Dim iFactura As Long
        Dim fCredito As Double
        Dim fCantPagar As Double
        Dim fFinanciamiento As Double
        Dim fTotalPagar As Double
        Dim lRow As Long
        Dim iCliente As Integer
        Dim iPagos As Integer
        
        For lRow = 1 To sprCreditos.DataRowCnt
            
            'Obtener los datos de la forma
            sprCreditos.Col = COL_CREDITOS_SELECCIONADO '1
            sprCreditos.Row = lRow
            
            If sprCreditos.Text = 1 Then
            
                sprCreditos.Col = COL_CREDITOS_CLIENTE '2
                iCliente = Val(sprCreditos.Text)
                
                sprCreditos.Col = COL_CREDITOS_CREDITO '4
                fCredito = Val(fnstrValor(sprCreditos.Text))
                
                sprCreditos.Col = COL_CREDITOS_CANT_PAGAR '8
                fCantPagar = Val(fnstrValor(sprCreditos.Text))
                
                sprCreditos.Col = COL_CREDITOS_FINANCIAMIENTO '6
                fFinanciamiento = Val(fnstrValor(sprCreditos.Text))
                
                sprCreditos.Col = COL_CREDITOS_TOTAL_PAGAR '9
                fTotalPagar = Val(fnstrValor(sprCreditos.Text))
                
                sprCreditos.Col = COL_CREDITOS_NO_PAGOS '7
                iPagos = Val(sprCreditos.Text)
                
                'Registrar el nuevo crédito, enviando los datos del credito.
                iFactura = oCredito.registraCredito(iCliente, _
                                                    fCredito, _
                                                    fCantPagar, _
                                                    fFinanciamiento, _
                                                    fTotalPagar, _
                                                    iPagos, _
                                                    Format(Now() + 1, "dd/mm/yyyy"), _
                                                    Format(Now + iPagos, "dd/mm/yyyy"), _
                                                    Format(Now(), "dd/mm/yyyy"), _
                                                    "V", _
                                                    "0", _
                                                    cbCobrador.Text, _
                                                    0, _
                                                    "")
                'Enviar el reporte de la factura a impresora, ejecutando la función privada 'imprimefn'
                Call imprimefn(iCliente, iFactura)
                'Enviar mensaje indicando que ya quedó el crédito registrado
                MsgBox "Ha quedado registrado el nuevo crédito", vbInformation + vbOKOnly
            End If
            
        Next lRow
        
        Set oCredito = Nothing
        
    Else
        'Si no fue aceptado (si la propiedad bPermiteAcceso de la ventana 'accesofrm' es = false)
        If iVeces >= 3 Then
            'Terminar y cerrar la pantalla
            Unload Me
        End If
    End If

End Sub

Private Function imprimefn(iCliente As Integer, iFactura As Long)

    'Declarar el uso de la clase Reporte
    'Definir las propiedades siguientes:
       ' EL objeto crystal report (el de la forma) sobre el cual se realizará el reporte
       ' Definir si el reporte puede ser a pantalla o directo a impresora
       ' Definir a que impresora se enviará el reporte
       ' Definir el nombre del reporte (el que se diseña para la impresión en el crystal reports)
       ' Definir los parámetros iCliente e iFactura (en una colección)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    cParametros.Add oCampo.CreaCampo(adInteger, , , iFactura)  'Factura
    
    oReporte.oCrystalReport = Me.crFactura
    If gstrReporteEnPantalla = "Si" Then
        oReporte.bVistaPreliminar = True
    Else
        oReporte.bVistaPreliminar = False
    End If
    oReporte.strImpresora = gPrintPed
    oReporte.strNombreReporte = DirSys & "factura.rpt"
    oReporte.cParametros = cParametros
    oReporte.fnImprime
    Set oReporte = Nothing

'COnsideraciones:

    'Utilizar la clase Reporte
    'antes de enviar ejecutar la impresión definir las propiedades siguientes:
        'oCrystalReport
        'bVistaPreliminar
        'strImpresora
        'strNombreReporte
        'cParametros
    'Ejectuar el reporte con la función fnImprime
    'hacer el objeto reporte igual a nothing.
    
End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub


