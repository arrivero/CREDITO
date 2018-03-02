VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reGeneralfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen Diario General"
   ClientHeight    =   5355
   ClientLeft      =   4455
   ClientTop       =   3270
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdevo 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   14
      Top             =   4230
      Width           =   1185
   End
   Begin VB.TextBox txtcheque 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   13
      Top             =   3870
      Width           =   1185
   End
   Begin VB.TextBox txtefeche 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Height          =   330
      Left            =   2925
      TabIndex        =   12
      Top             =   3510
      Width           =   1185
   End
   Begin VB.CommandButton cmdCorte 
      Caption         =   "Corte y Reporte"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin FPSpread.vaSpread sprFaltanteSobrante 
      Height          =   1725
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   5175
      _Version        =   196608
      _ExtentX        =   9128
      _ExtentY        =   3043
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
      MaxCols         =   5
      MaxRows         =   6
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "reGeneralfrm.frx":0000
   End
   Begin Crystal.CrystalReport crReporteGeneral 
      Left            =   210
      Top             =   4860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   4770
      Width           =   1095
   End
   Begin VB.CommandButton cmdreporte 
      Caption         =   "Generar Reporte"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   4770
      Width           =   1095
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtfecha 
      Height          =   345
      Left            =   1920
      TabIndex        =   6
      Top             =   450
      Width           =   1965
      _Version        =   65537
      _ExtentX        =   3466
      _ExtentY        =   609
      _StockProps     =   93
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin EditLib.fpCurrency txtefecheold 
      Height          =   315
      Left            =   2910
      TabIndex        =   7
      Top             =   3510
      Visible         =   0   'False
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
      _ExtentY        =   556
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtchequeold 
      Height          =   315
      Left            =   2910
      TabIndex        =   8
      Top             =   3870
      Visible         =   0   'False
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
      _ExtentY        =   556
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtdevoold 
      Height          =   315
      Left            =   2910
      TabIndex        =   9
      Top             =   4230
      Visible         =   0   'False
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
      _ExtentY        =   556
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label5 
      Caption         =   "Devoluciones:"
      Height          =   255
      Left            =   1470
      TabIndex        =   5
      Top             =   4275
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Cheque:"
      Height          =   255
      Left            =   1470
      TabIndex        =   4
      Top             =   3915
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Efectivo:"
      Height          =   255
      Left            =   1470
      TabIndex        =   3
      Top             =   3555
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5310
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   1140
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   1020
      Y2              =   1020
   End
End
Attribute VB_Name = "reGeneralfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nombreglobal As String
Dim bgral As Integer

Private Const COL_CHK = 1
Private Const COL_TIPO = 2
Private Const COL_MONTO = 3
Private Const COL_COBRADOR = 4
Private Const COL_ID_COBRADOR = 5

Private Sub txtfecha_Click()

    Dim dEfectivo As Double
    Dim dCheques As Double
    Dim dDevolucion As Double
    Dim cSobrantes As New Collection
    Dim oPago As New Pago

    'Obten el efectivo, cheques y devolucion
    oPago.obtenResumenDiarioMontos dEfectivo, dCheques, dDevolucion, txtfecha.Text
    txtefeche = dEfectivo
    txtcheque = dCheques
    txtdevo = dDevolucion

    Set cSobrantes = oPago.obtenSobrantes(txtfecha.Text)
    
    If cSobrantes.Count > 0 Then
        fnLlenaTablaCollection sprFaltanteSobrante, cSobrantes
    End If

    Set oPago = Nothing

End Sub

Private Sub Form_Load()
    
    Dim dEfectivo As Double
    Dim dCheques As Double
    Dim dDevolucion As Double
    
    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        llenaComboSpread Me.sprFaltanteSobrante, COL_COBRADOR, oUsuario.cDatos, 0
    End If
    Set oUsuario = Nothing
        
    Dim oPago As New Pago
    Dim cSobrantes As New Collection
     
    Set cSobrantes = oPago.obtenSobrantes(txtfecha.Text)
    
    If cSobrantes.Count > 0 Then
        fnLlenaTablaCollection sprFaltanteSobrante, cSobrantes
    End If
    
    'Obten el efectivo, cheques y devolucion
    oPago.obtenResumenDiarioMontos dEfectivo, dCheques, dDevolucion, txtfecha.Text
    txtefeche = dEfectivo
    txtcheque = dCheques
    txtdevo = dDevolucion
    
    Set oPago = Nothing

    
'    txtfecha.Text = Format(Now, "dd/mm/yyyy")
'    txtcantidad.Text = 0#
'
'    Dim suma As Double
'    Dim datos As Recordset
'    Dim fecha As Date
'
'    fecha = Format(Now, "dd/mm/yyyy")
'    Set datos = Base.OpenRecordset("Select * from Sobrantes where cdbl(fecha) = " + CStr(CDbl(fecha)))
'    suma = 0
'
'    While Not datos.EOF
'        If CStr(datos!tipo) = "S" Then
'            suma = suma + CDbl(datos!cantidad)
'        Else
'            If CStr(datos!tipo) = "F" Then
'                suma = suma - CDbl(datos!cantidad)
'            End If
'        End If
'        datos.MoveNext
'    Wend
'
'    If suma < 0 Then
'        optfaltante.Value = True
'        suma = suma - (2 * suma)
'    Else
'        If suma >= 0 Then
'            optsobrante.Value = True
'        End If
'    End If
'
'    txtcantidad.Text = suma

End Sub

Private Sub fnImprime(strReporte As String, crObjeto As CrystalReport, strFechaInicial As String)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    Dim strNombreReporte As String
    
'    If strFechaInicial <> "" Then

        cParametros.Add oCampo.CreaCampo(adInteger, , , 1)

'    End If
    
    strNombreReporte = strReporte + ".rpt"
    
    oReporte.oCrystalReport = crObjeto
    If gstrReporteEnPantalla = "Si" Then
        oReporte.bVistaPreliminar = True
    Else
        oReporte.bVistaPreliminar = False
    End If
    
    oReporte.strImpresora = gPrintPed
    oReporte.strNombreReporte = DirSys & strNombreReporte
    oReporte.cParametros = cParametros
    oReporte.fnImprime
    Set oReporte = Nothing
    
End Sub

Private Function obtenFaltantesSobrantes() As Collection
    
    Dim cRegistros As New Collection
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    
    Dim lCol, lRow As Long
    Dim fMonto As Double
    
    For lRow = 1 To sprFaltanteSobrante.DataRowCnt
    
        sprFaltanteSobrante.Row = lRow
        sprFaltanteSobrante.Col = COL_MONTO
        
        fMonto = Val(fnstrValor(sprFaltanteSobrante.Text))
        
        If fMonto > 0 Then
            
            Set cRegistro = New Collection
            
            sprFaltanteSobrante.Col = COL_CHK
            
            cRegistro.Add oCampo.CreaCampo(adInteger, , , txtfecha)
            
            If sprFaltanteSobrante.Text = 1 Then
                cRegistro.Add oCampo.CreaCampo(adInteger, , , "F")
            Else
                cRegistro.Add oCampo.CreaCampo(adInteger, , , "S")
            End If
            
            cRegistro.Add oCampo.CreaCampo(adInteger, , , fMonto)
            sprFaltanteSobrante.Col = COL_COBRADOR
            cRegistro.Add oCampo.CreaCampo(adInteger, , , sprFaltanteSobrante.Text)
            
            cRegistros.Add cRegistro
            
        End If
                
    Next lRow
    
    Set obtenFaltantesSobrantes = cRegistros
    
End Function

Private Sub cmdCorte_Click()

    Dim strCobradorDescansa As String
    Dim strCobradorSuple As String
    
    cambioCobrador.Show vbModal
    
    If cambioCobrador.bCancelado = False Then

        strCobradorDescansa = cambioCobrador.strCobradorDescansa
        strCobradorSuple = cambioCobrador.strCobradorSuple
        
        'Poner el cursor de proceso
        Screen.MousePointer = vbHourglass
    
        Dim oPago As New Pago
        
        'fnLlenaTablaCollection sprFaltanteSobrante, oPago.obtenSobrantes(txtfecha.Text)
        
        If DateDiff("d", Date, txtfecha.Text) = 0 Then
            'registra faltantes y sobrantes
            Call oPago.registraSobrante(obtenFaltantesSobrantes, txtfecha.Text)
        Else
            
            sprFaltanteSobrante.Col = -1
            sprFaltanteSobrante.Col2 = -1
            sprFaltanteSobrante.Row = -1
            sprFaltanteSobrante.Row2 = -1
            sprFaltanteSobrante.BlockMode = True
            sprFaltanteSobrante.Lock = True
            sprFaltanteSobrante.BlockMode = False
            
            'fnLlenaTablaCollection sprFaltanteSobrante, oPago.obtenSobrantes(txtfecha.Text)
        End If
        Set oPago = Nothing
        
        'PREPARA EL CORTE - RESUMEN GENERAL PARA EL REPORTE
        Dim oCredito As New credito
        oCredito.corte Val(fnstrValor(txtefeche.Text)), Val(fnstrValor(txtcheque.Text)), Val(fnstrValor(txtdevo.Text)), txtfecha.Text, 1, strCobradorDescansa, strCobradorSuple
    
        Set oCredito = Nothing
        
        'GENERA EL REPORTE
        Call fnImprime("reporteGeneral", crReporteGeneral, txtfecha.Text)
    
    Else
        MsgBox "El corte no se ejecutará", vbInformation + vbOKOnly
    End If
    
    'Poner el cursor Normal
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdreporte_Click()

    
    'Poner el cursor de proceso
    Screen.MousePointer = vbHourglass
    
    Dim oPago As New Pago
    
    'fnLlenaTablaCollection sprFaltanteSobrante, oPago.obtenSobrantes(txtfecha.Text)
    
    If DateDiff("d", Date, txtfecha.Text) = 0 Then
        'registra faltantes y sobrantes
        Call oPago.registraSobrante(obtenFaltantesSobrantes, txtfecha.Text)
    Else
        
        
        Dim dEfectivo As Double
        Dim dCheques As Double
        Dim dDevolucion As Double
            
        'Obten el efectivo, cheques y devolucion
        oPago.obtenResumenDiarioMontos dEfectivo, dCheques, dDevolucion, txtfecha.Text
        txtefeche = dEfectivo
        txtcheque = dCheques
        txtdevo = dDevolucion
        
        sprFaltanteSobrante.Col = -1
        sprFaltanteSobrante.Col2 = -1
        sprFaltanteSobrante.Row = -1
        sprFaltanteSobrante.Row2 = -1
        sprFaltanteSobrante.BlockMode = True
        sprFaltanteSobrante.Lock = True
        sprFaltanteSobrante.BlockMode = False
        
        fnLlenaTablaCollection sprFaltanteSobrante, oPago.obtenSobrantes(txtfecha.Text)
        
        
    End If
    
    Set oPago = Nothing
    
    'PREPARA EL CORTE - RESUMEN GENERAL PARA EL REPORTE
    Dim oCredito As New credito
    oCredito.corte Val(fnstrValor(txtefeche.Text)), Val(fnstrValor(txtcheque.Text)), Val(fnstrValor(txtdevo.Text)), txtfecha.Text, 0, "", ""
    Set oCredito = Nothing
    
    'GENERA EL REPORTE
    Call fnImprime("reporteGeneral", crReporteGeneral, txtfecha.Text)

    
    'Poner el cursor Normal
    Screen.MousePointer = vbDefault

End Sub

Private Sub sprFaltanteSobrante_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    
    sprFaltanteSobrante.Col = COL_CHK
    sprFaltanteSobrante.Row = sprFaltanteSobrante.ActiveRow
    If sprFaltanteSobrante.Text = 1 Then
        sprFaltanteSobrante.Col = COL_TIPO
        sprFaltanteSobrante.Text = "Faltante"
    Else
        sprFaltanteSobrante.Col = COL_TIPO
        sprFaltanteSobrante.Text = "Sobrante"
    End If
        
    sprFaltanteSobrante.Col = COL_MONTO
    sprFaltanteSobrante.Action = ActionActiveCell
        
End Sub

Private Sub sprFaltanteSobrante_KeyPress(KeyAscii As Integer)

    If vbKeyReturn = KeyAscii Then
    
            Select Case sprFaltanteSobrante.ActiveCol
                Case Is = COL_MONTO
                    sprFaltanteSobrante.Col = COL_COBRADOR
                    sprFaltanteSobrante.Action = ActionActiveCell
                Case Is = COL_COBRADOR
                    sprFaltanteSobrante.Col = COL_CHK
                    sprFaltanteSobrante.Row = sprFaltanteSobrante.ActiveRow + 1
                    sprFaltanteSobrante.Action = ActionActiveCell

            End Select
            
    End If
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

'Private Sub cd()
'
'Dim i, j As Integer
'Dim datos As Recordset
'Dim folio1, folio2, folio3 As Long
'Dim pago1, pago2, pago3 As Double
'Dim fecha As Date
'
'folio1 = 0
'folio2 = 0
'folio3 = 0
'
'pago1 = 0
'pago2 = 0
'pago3 = 0
'
'j = O
'
''base.Execute "delete from cortediario"
'
'Set datos = Base.OpenRecordset("select * from cdiario")
''datos.MoveFirst
'While Not datos.EOF
'    For i = 1 To 3
'        If datos.EOF Then
'            GoTo fincd
'        Else
'            Select Case i
'            Case 1
'                folio1 = datos!factura
'                pago1 = datos!Cantpagada
'            Case 2
'                folio2 = datos!factura
'                pago2 = datos!Cantpagada
'            Case 3
'                folio3 = datos!factura
'                pago3 = datos!Cantpagada
'            End Select
'        End If
'        fecha = datos!fecha
'        datos.MoveNext
'    Next i
'fincd:
'    Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,fechacd,folio1,pago1,folio2,pago2,folio3,pago3) values(" + CStr(1) + "," + CStr(1) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + ",'" + CStr(fecha) + "'," + CStr(folio1) + "," + CStr(pago1) + "," + CStr(folio2) + "," + CStr(pago2) + "," + CStr(folio3) + "," + CStr(pago3) + ")"
'    j = j + 1
'    folio1 = 0
'    folio2 = 0
'    folio3 = 0
'    pago1 = 0
'    pago2 = 0
'    pago3 = 0
'Wend
'
'datos.Close
'
''If j > 0 Then
''    diario.PrintReport
''Else
''    MsgBox "No existen pagos registrados para el dia de hoy", vbInformation, "Corte Diario"
''End If
'
'End Sub
'
'
'Private Sub nuevos()
'Dim datos, datos1, datos2 As Recordset
'Dim pagos1, gastos, folio As Double
'Dim conta As Integer
'
'If txtfecha.Text <> "" Then
'
'    conta = 0
'    Set datos = Base.OpenRecordset("select * from qrycreditos")
'    While Not datos.EOF
'        If datos!fecha = CDate(txtfecha.Text) Then
'            Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fechac,Status,elec) values(" + CStr(2) + "," + CStr(0) + "," + CStr(2) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!credito) + ",'" + CStr(datos!Nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "'," + CStr(IIf(IsNull(datos!electrico), 0, datos!electrico)) + ")"
'            conta = 1
'        End If
'        datos.MoveNext
'    Wend
'    datos.Close
'Else
'End If
'End Sub
'
'
'Private Sub termina()
'Dim datos, datos1, datos2 As Recordset
'Dim pagos1, gastos, folio As Double
'Dim conta As Integer
'
'If txtfecha.Text <> "" Then
'
'    conta = 0
'    Set datos = Base.OpenRecordset("select * from qrycreditosterminados ")
'    While Not datos.EOF
'        If datos!fechatermina = CDate(txtfecha.Text) Then
'            Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fechac,Status) values(" + CStr(3) + "," + CStr(0) + "," + CStr(0) + "," + CStr(3) + "," + CStr(0) + "," + CStr(0) + "," + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!credito) + ",'" + CStr(datos!Nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "')"
'            conta = 1
'        End If
'        datos.MoveNext
'    Wend
'    datos.Close
'
'Else
'End If
'End Sub
'
'
'
'
'Private Sub gastos()
'
'Dim datos, datos1, datos2 As Recordset
'Dim pagos1, gastos, folio As Double
'Dim conta As Integer
'
'If txtfecha.Text <> "" Then
'
'    conta = 0
'    Set datos = Base.OpenRecordset("select * from gastos_dia")
'    While Not datos.EOF
'        If datos!fecha = CDate(txtfecha.Text) Then
'            Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,gasto,fechag,importe,descripcion) values(" + CStr(4) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(4) + "," + CStr(0) + "," + CStr(datos!Gasto) + ",'" + Format(datos!fecha, "dd/mm/yyyy") + "','" + CStr(datos!importe) + "','" + datos!descripcion + "')"
'            conta = 1
'        End If
'        datos.MoveNext
'    Wend
'    datos.Close
'
'Else
'End If
'
'
'End Sub
'
'
'
'Private Sub resumen()
'
'Dim datos, datos1, datos2 As Recordset
'Dim pagos1, gastos, folio, folioe, total_anterior As Double
'Dim conta As Integer
'
'If txtfecha.Text <> "" Then
'
'    total_anterior = 0
'    conta = 0
'    Set datos2 = Base.OpenRecordset("select * from total_dia_anterior")
'    While Not datos2.EOF
'        'If DatePart("d", datos2!fecha) = DatePart("d", CDate(txtfecha.Text)) And DatePart("m", datos2!fecha) = DatePart("m", CDate(txtfecha.Text)) And DatePart("yyyy", datos2!fecha) = DatePart("yyyy", CDate(txtfecha.Text)) Then
'        If datos2!fecha = CDate(txtfecha.Text) Then
'            datos2.MovePrevious
'            conta = 1
'        End If
'
'        total_anterior = IIf(IsNull(datos2!total_anterior), 0, datos2!total_anterior)
'        If conta = 1 Then
'            datos2.MoveLast
'            datos2.MoveNext
'        Else
'            datos2.MoveNext
'        End If
'    Wend
'    datos2.Close
'
'
'    '*********************************************************************
'    If conta = 0 Then
'        Set datos2 = Base.OpenRecordset("select * from total_dia_anterior where clng(fecha) < " + CStr(CLng(CDate(txtfecha.Text))) + " order by fecha desc")
'        While Not datos2.EOF
'            total_anterior = IIf(IsNull(datos2!total_anterior), 0, datos2!total_anterior)
'            datos2.MoveLast
'            datos2.MoveNext
'        Wend
'        datos2.Close
'    End If
'    '*********************************************************************
'
'
'    Base.Execute "delete from resumen_diario_gral where cstr(datepart('d',fecha))=" + CStr(DatePart("d", CDate(txtfecha.Text))) + " and cstr(datepart('m',fecha))=" + CStr(DatePart("m", CDate(txtfecha.Text))) + " and cstr(datepart('yyyy',fecha))=" + CStr(DatePart("yyyy", CDate(txtfecha.Text)))
'    'base.Execute "delete from resumen_diario"
'
'    conta = 0
'    Set datos = Base.OpenRecordset("select * from qrysumapagos")
'    While Not datos.EOF
'        If datos!fecha = CDate(txtfecha.Text) Then
'            pagos1 = datos!pagos
'            conta = 1
'        End If
'        datos.MoveNext
'    Wend
'    datos.Close
'    If conta = 0 Then
'        pagos1 = 0
'    End If
'
'    conta = 0
'    Set datos1 = Base.OpenRecordset("select * from qrysumacredito")
'    While Not datos1.EOF
'        If datos1!fecha = CDate(txtfecha.Text) Then
'            folio = datos1!credito
'            conta = 1
'        End If
'        datos1.MoveNext
'    Wend
'    datos1.Close
'    If conta = 0 Then
'        folio = 0
'    End If
'
'    conta = 0
'    Set datos1 = Base.OpenRecordset("select * from qrysumacreditoelectricos")
'    While Not datos1.EOF
'        If datos1!fecha = CDate(txtfecha.Text) Then
'            folioe = datos1!creditoe
'            conta = 1
'        End If
'        datos1.MoveNext
'    Wend
'    datos1.Close
'    If conta = 0 Then
'        folioe = 0
'    End If
'
'    conta = 0
'    Set datos2 = Base.OpenRecordset("select * from qrysumagastos")
'    While Not datos2.EOF
'        If datos2!fecha = CDate(txtfecha.Text) Then
'            gastos = datos2!importe
'            conta = 1
'        End If
'        datos2.MoveNext
'    Wend
'    datos2.Close
'    If conta = 0 Then
'        gastos = 0
'    End If
'
'
'    If txtcantidad.Text = "" Then
'        txtcantidad.Text = 0
'        If Not IsNumeric(txtcantidad.Text) Then
'            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
'            txtcantidad.Text = 0
'            txtcantidad.SetFocus
'        End If
'    End If
'    If Me.txtefeche.Text = "" Then
'        txtefeche.Text = 0
'        If Not IsNumeric(txtefeche.Text) Then
'            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
'            txtefeche.Text = 0
'            txtefeche.SetFocus
'        End If
'    End If
'
'    If Me.txtcheque.Text = "" Then
'        txtcheque.Text = 0
'        If Not IsNumeric(txtcheque.Text) Then
'            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
'            txtcheque.Text = 0
'            txtcheque.SetFocus
'        End If
'    End If
'
'    If Me.txtdevo.Text = "" Then
'        txtdevo.Text = 0
'        If Not IsNumeric(txtdevo.Text) Then
'            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
'            txtdevo.Text = 0
'            txtdevo.SetFocus
'        End If
'    End If
'
'    If optsobrante.Value = True Then
'        Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,fechard,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,electricos) values(" + CStr(5) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(0) + "," + CStr(txtcantidad.Text) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(folioe) + ")"
'        'base.Execute "insert into regen (cve,fechard,cobrado,facturas,gastos,faltante,sobrante) values(" + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(0) + "," + CStr(txtcantidad.Text) + ")"
'        Base.Execute "insert into resumen_diario_gral (fecha,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,anterior) values('" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(0) + "," + CStr(txtcantidad.Text) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(total_anterior) + ")"
'    Else
'        Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,fechard,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,electricos) values(" + CStr(5) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(txtcantidad.Text) + "," + CStr(0) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(folioe) + ")"
'        'base.Execute "insert into regen (cve,fechard,cobrado,facturas,gastos,faltante,sobrante) values(" + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(txtcantidad.Text) + "," + CStr(0) + ")"
'        Base.Execute "insert into resumen_diario_gral (fecha,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,anterior) values('" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(txtcantidad.Text) + "," + CStr(0) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(total_anterior) + ")"
'    End If
'    'rediario.PrintReport
'
'
'    total_anterior = 0
'    conta = 0
'    'Set datos2 = base.OpenRecordset("select * from resumen_por_fecha where DatePart('d', fecha) >= " + CStr(DatePart("d", CDate(txtfecha.Text))) + " And DatePart('m', fecha) >= " + CStr(DatePart("m", CDate(txtfecha.Text))) + " And DatePart('yyyy', fecha) >= " + CStr(DatePart("yyyy", CDate(txtfecha.Text))) + " order by fecha asc")
'    Set datos2 = Base.OpenRecordset("select * from resumen_por_fecha where clng(fecha) >= " + CStr(CLng(CDate(txtfecha.Text))) + " order by fecha asc")
'    While Not datos2.EOF
'        'If DatePart("d", datos2!fecha) = DatePart("d", CDate(txtfecha.Text)) And DatePart("m", datos2!fecha) = DatePart("m", CDate(txtfecha.Text)) And DatePart("yyyy", datos2!fecha) = DatePart("yyyy", CDate(txtfecha.Text)) Then
'        If conta = 1 Then
'            Base.Execute "update resumen_diario_gral set anterior=" + CStr(total_anterior) + " where datepart('d',fecha)=" + CStr(DatePart("d", datos2!fecha)) + " and datepart('m',fecha)=" + CStr(DatePart("m", datos2!fecha)) + " and datepart('yyyy',fecha)=" + CStr(DatePart("yyyy", datos2!fecha))
'        End If
'        If datos2!fecha = CDate(txtfecha.Text) Then
'            conta = 1
'            total_anterior = (((((((datos2!anterior_ + datos2!cobrado) - datos2!facturas) - datos2!gastos) - datos2!faltante) + datos2!sobrante) + datos2!cheques) - datos2!depositos) - datos2!gtos + datos2!devo + datos2!electricos
'        Else
'            total_anterior = (((((((total_anterior + datos2!cobrado) - datos2!facturas) - datos2!gastos) - datos2!faltante) + datos2!sobrante) + datos2!cheques) - datos2!depositos) - datos2!gtos + datos2!devo + datos2!electricos
'        End If
'        datos2.MoveNext
'    Wend
'    datos2.Close
'
'Else
''    MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario"
''    txtfecha.Text = Format(Now, "dd/mm/yyyy")
''    txtfecha.SetFocus
'End If
'
'
'
'End Sub


'Private Sub repgeneral()
'
'Dim xlaCPP As Object 'Aplicacion
'Dim xlwCPP As Object 'archivo
'Dim xlsCPP As Object 'Worksheet
'
'Dim bdd As DataBase
'Dim r As Recordset
'Dim a, b, FechaInicio, archivo As String
'Dim i, j, k, Fila, Columna As Integer
'Dim res As Integer
'Dim cdtotal As Double
'
'archivo = "c:\facturas\Reporte\General"
'''archivo = archivo + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")" + ".xls"
''Set bdd = CurrentDb
''Set xlwCPP = CreateObject("Excel.Sheet.8")
''Set xlwCPP = GetObject(archivo)
''xlwCPP.Application.Visible = True
''xlwCPP.Parent.Windows(1).Visible = True
'
''Crear el archivo de Excel
'Set xlwCPP = CreateObject("Excel.Sheet.8")
'Set xlsCPP = xlwCPP.Activesheet
'Set xlaCPP = xlsCPP.Parent.Parent
''xlwCPP.SaveAs (archivo)
'xlwCPP.Application.Visible = True
'xlaCPP.ActiveWindow.WindowState = -4137 'Maximiza la ventana
'
''xlsCPP.Name = "General"
'xlsCPP.Name = "General" + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")"
'
''''''''''''''''''''''''''''''''''''''''''' Resumen
'Fila = 5
'Columna = 1
'j = 5
'Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 5")
'Set xlsCPP = xlwCPP.Worksheets("General" + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")")
'xlwCPP.Worksheets("General" + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")").Activate
'Set xlsCPP = xlwCPP.Activesheet
'
'
'xlaCPP.Cells(1, 1).Value = "Resumen Diario"
'xlaCPP.Range("A1:B1").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 12
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(1, 4).Value = "Fecha: "
'xlaCPP.Range("D1").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = False
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 12
'xlaCPP.Selection.Font.Bold = True
'
'
'xlaCPP.Cells(2, 4).Value = "Total Cobrado:"
'xlaCPP.Range("D2:E2").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
'xlaCPP.Cells(3, 4).Value = "'-Facturas:"
'xlaCPP.Range("D3:E3").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
'xlaCPP.Cells(4, 4).Value = "'-Gastos:"
'xlaCPP.Range("D4:E4").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
''xlaCPP.Cells(5, 4).Value = "Total:"
'xlaCPP.Cells(5, 4).Value = "+Electrónicos:"
'xlaCPP.Range("D5:E5").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
'xlaCPP.Cells(6, 4).Value = "'-Faltante:"
'xlaCPP.Range("D6:E6").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
'xlaCPP.Cells(7, 4).Value = "'+Sobrante:"
'xlaCPP.Range("D7:E7").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
'xlaCPP.Cells(8, 4).Value = "+Efectivo:"
'xlaCPP.Range("D8:E8").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
''*************************total*********************************************
'xlaCPP.Cells(9, 4).Value = "+Cheque:"
'xlaCPP.Range("D9:E9").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
''***********************************************************************
'xlaCPP.Cells(10, 4).Value = "+Devoluciones:"
'xlaCPP.Range("D10:E10").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
'xlaCPP.Cells(11, 4).Value = "Total:"
'xlaCPP.Range("D11:E11").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = False
'
''xlaCPP.Range("F4").Select
''    With xlaCPP.Selection.Borders(4)
''        .LineStyle = xlContinuous
''        .Weight = 3
''        .ColorIndex = xlAutomatic
''    End With
'
'xlaCPP.Range("F10").Select
'    With xlaCPP.Selection.Borders(4)
'        .LineStyle = xlContinuous
'        .Weight = 3
'        .ColorIndex = xlAutomatic
'    End With
'
'xlaCPP.Range("A11:I11").Select
'    With xlaCPP.Selection.Borders(4)
'        .LineStyle = xlContinuous
'        .Weight = 3
'        .ColorIndex = xlAutomatic
'    End With
'
'While Not r.EOF
'    xlaCPP.Cells(1, 5).Value = r!fechard
'
'    xlaCPP.Range("E1:F1").Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.MergeCells = True
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 11
'    xlaCPP.Selection.Font.Bold = True
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'
'    xlaCPP.Cells(2, 6).Value = Format(r!cobrado, "###,###,###,###0.00")
'    xlaCPP.Cells(3, 6).Value = Format(r!facturas, "###,###,###,###0.00")
'    xlaCPP.Cells(4, 6).Value = Format(r!gastos, "###,###,###,###0.00")
'    'xlaCPP.Cells(5, 6).Value = Format((r!cobrado - r!facturas) - r!gastos, "###,###,###,###0.00")
'    xlaCPP.Cells(5, 6).Value = Format(r!electricos, "###,###,###,###0.00")
'    xlaCPP.Cells(6, 6).Value = Format(r!faltante, "###,###,###,###0.00")
'    xlaCPP.Cells(7, 6).Value = Format(r!sobrante, "###,###,###,###0.00")
'    xlaCPP.Cells(8, 6).Value = Format(r!ec, "###,###,###,###0.00")
'    xlaCPP.Cells(9, 6).Value = Format(r!cheque, "###,###,###,###0.00")
'    xlaCPP.Cells(10, 6).Value = Format(r!devolucion, "###,###,###,###0.00")
'    xlaCPP.Cells(11, 6).Value = Format((((r!cobrado - r!facturas) - r!gastos) - r!faltante) + r!sobrante + r!ec + r!cheque + r!devolucion + r!electricos, "###,###,###,###0.00")
'    j = j + 1
'    r.MoveNext
'Wend
'r.Close
'
'''''''''''''''''''''''''''''''''''''''''''CORTE DIARIO
'
'xlaCPP.Cells(13, 1).Value = "Corte Diario"
'xlaCPP.Range("A13:B13").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 12
'xlaCPP.Selection.Font.Bold = True
'
'a = "B"
'b = "C"
'j = 2
'For i = 1 To 3
'    xlaCPP.Cells(15, j).Value = "Folio"
'    xlaCPP.Range(a + "15").Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.MergeCells = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = True
'
'    xlaCPP.Cells(15, j + 1).Value = "Cantidad"
'    xlaCPP.Range(b + "15").Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.MergeCells = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = True
'    j = j + 3
'    If i = 1 Then
'        a = "E"
'        b = "F"
'    Else
'        a = "H"
'        b = "I"
'    End If
'Next i
'
'xlaCPP.Cells(13, 4).Value = "Fecha:"
'xlaCPP.Range("D13").Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = False
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = True
'
'j = 16
'cdtotal = 0
'i = 1
'Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 1")
'If Not r.EOF Then
'    xlaCPP.Cells(13, 5).Value = IIf(IsNull(r!fechacd), " ", r!fechacd)
'    xlaCPP.Range("E13").Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = True
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'End If
'
'While Not r.EOF
'
'    xlaCPP.Cells(j, 2).Value = r!folio1
'    xlaCPP.Cells(j, 3).Value = Format(r!pago1, "###,###,###,###0.00")
'    xlaCPP.Cells(j, 5).Value = r!folio2
'    xlaCPP.Cells(j, 6).Value = Format(r!pago2, "###,###,###,###0.00")
'    xlaCPP.Cells(j, 8).Value = r!folio3
'    xlaCPP.Cells(j, 9).Value = Format(r!pago3, "###,###,###,###0.00")
'
'    j = j + 1
'    i = i + 1
'    cdtotal = cdtotal + r!pago1 + r!pago2 + r!pago3
'    r.MoveNext
'    If i = 20 And Not r.EOF Then
'        a = "B"
'        b = "C"
'        k = 2
'        For i = 1 To 3
'            xlaCPP.Cells(j, k).Value = "Folio"
'            xlaCPP.Range(a + CStr(j)).Select
'            xlaCPP.Selection.WrapText = False
'            xlaCPP.Selection.Orientation = 0
'            xlaCPP.Selection.AddIndent = False
'            xlaCPP.Selection.MergeCells = False
'            xlaCPP.Selection.Font.Name = "Arial"
'            xlaCPP.Selection.Font.Size = 10
'            xlaCPP.Selection.Font.Bold = True
'
'            xlaCPP.Cells(j, k + 1).Value = "Cantidad"
'            xlaCPP.Range(b + CStr(j)).Select
'            xlaCPP.Selection.WrapText = False
'            xlaCPP.Selection.Orientation = 0
'            xlaCPP.Selection.AddIndent = False
'            xlaCPP.Selection.MergeCells = False
'            xlaCPP.Selection.Font.Name = "Arial"
'            xlaCPP.Selection.Font.Size = 10
'            xlaCPP.Selection.Font.Bold = True
'            k = k + 3
'            If i = 1 Then
'                a = "E"
'                b = "F"
'            Else
'                a = "H"
'                b = "I"
'            End If
'        Next i
'        j = j + 1
'        i = 1
'    End If
'
'Wend
'r.Close
'
'a = "B"
'b = "C"
'For i = 1 To 3
'    xlaCPP.Range(a + "15:" + b + CStr(j - 1)).Select
'        With xlaCPP.Selection.Borders(1)
'            .LineStyle = xlContinuous
'            .Weight = 2
'            .ColorIndex = xlAutomatic
'        End With
'        With xlaCPP.Selection.Borders(2)
'            .LineStyle = xlContinuous
'            .Weight = 2
'            .ColorIndex = xlAutomatic
'        End With
'        With xlaCPP.Selection.Borders(4)
'            .LineStyle = xlContinuous
'            .Weight = 2
'            .ColorIndex = xlAutomatic
'        End With
'        With xlaCPP.Selection.Borders(3)
'            .LineStyle = xlContinuous
'            .Weight = 2
'            .ColorIndex = xlAutomatic
'        End With
'    If i = 1 Then
'        a = "E"
'        b = "F"
'    Else
'        a = "H"
'        b = "I"
'    End If
'Next i
'
'xlaCPP.Cells(j + 1, 6).Value = "Total Corte Diario"
'xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
''xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
'xlaCPP.Range("H" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'j = j + 2
'
'xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
'    With xlaCPP.Selection.Borders(4)
'        .LineStyle = xlContinuous
'        .Weight = 3
'        .ColorIndex = xlAutomatic
'    End With
'
'j = j + 2
'
'''''''''''''''''''''''''''''''''''''''''''CRÉDITOS NUEVOS
'
'xlaCPP.Cells(j, 1).Value = "Créditos Nuevos"
'xlaCPP.Range("A" + CStr(j) + ":B" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 12
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 4).Value = "Fecha Alta Crédito:"
'xlaCPP.Range("D" + CStr(j) + ":E" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = True
'
'Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 2")
'If Not r.EOF Then
'    xlaCPP.Cells(j, 6).Value = IIf(IsNull(r!Fechac), " ", r!Fechac)
'    xlaCPP.Range("F" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = True
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'End If
'j = j + 2
'
'xlaCPP.Cells(j, 1).Value = "       Folio"
'xlaCPP.Range("A" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 3).Value = "Nombre"
'xlaCPP.Range("C" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 6).Value = "      Crédito"
'xlaCPP.Range("F" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 8).Value = "     Inicio"
'xlaCPP.Range("H" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'
'xlaCPP.Cells(j, 9).Value = "       Fin"
'xlaCPP.Range("I" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'j = j + 1
'cdtotal = 0
'While Not r.EOF
'
'    xlaCPP.Cells(j, 1).Value = CStr(r!factura)
'    xlaCPP.Range("A" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 3).Value = r!Nombre
'    xlaCPP.Range("C" + CStr(j) + ":D" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 5).Value = IIf(r!elec = 1, "Electricos", "")
'    xlaCPP.Range("E" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 6).Value = Format(r!credito, "###,###,###,###0.00")
'    xlaCPP.Range("F" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 8).Value = r!fechaini
'    xlaCPP.Range("H" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'
'    xlaCPP.Cells(j, 9).Value = r!fechatermina
'    xlaCPP.Range("I" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'
'    j = j + 1
'    i = i + 1
'    cdtotal = cdtotal + r!credito
'    r.MoveNext
'    If i = 20 And Not r.EOF Then
'        a = "B"
'        b = "C"
'        k = 2
'        xlaCPP.Cells(j, 1).Value = "       Folio"
'        xlaCPP.Range("A" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 3).Value = "Nombre"
'        xlaCPP.Range("C" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 6).Value = "      Crédito"
'        xlaCPP.Range("F" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 8).Value = "     Inicio"
'        xlaCPP.Range("H" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = False
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 9).Value = "       Fin"
'        xlaCPP.Range("I" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        j = j + 1
'        i = 1
'    End If
'
'Wend
'r.Close
'
'xlaCPP.Cells(j + 1, 6).Value = "Total Créditos Nuevos"
'xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
''xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
'xlaCPP.Range("H" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'j = j + 2
'
'xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
'    With xlaCPP.Selection.Borders(4)
'        .LineStyle = xlContinuous
'        .Weight = 3
'        .ColorIndex = xlAutomatic
'    End With
'
'j = j + 2
'
'''''''''''''''''''''''''''''''''''''''''''CRÉDITOS TERMINADOS
'
'xlaCPP.Cells(j, 1).Value = "Créditos Terminados"
''xlaCPP.Range("A" + CStr(j) + ":B" + CStr(j)).Select
'xlaCPP.Range("A" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 12
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 4).Value = "Fecha Terminación Crédito:"
''xlaCPP.Range("D" + CStr(j) + ":E" + CStr(j)).Select
'xlaCPP.Range("D" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = True
'
'Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 3")
'If Not r.EOF Then
'    xlaCPP.Cells(j, 7).Value = IIf(IsNull(r!fechatermina), " ", r!fechatermina)
'    xlaCPP.Range("G" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = True
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'End If
'
'j = j + 2
'
'xlaCPP.Cells(j, 1).Value = "       Folio"
'xlaCPP.Range("A" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 3).Value = "Nombre"
'xlaCPP.Range("C" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 6).Value = "      Crédito"
'xlaCPP.Range("F" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 8).Value = "      Alta"
'xlaCPP.Range("H" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'
'xlaCPP.Cells(j, 9).Value = "     Inicio"
'xlaCPP.Range("I" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'j = j + 1
'cdtotal = 0
'While Not r.EOF
'
'    xlaCPP.Cells(j, 1).Value = CStr(r!factura)
'    xlaCPP.Range("A" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 3).Value = r!Nombre
'    xlaCPP.Range("C" + CStr(j) + ":E" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 6).Value = Format(r!credito, "###,###,###,###0.00")
'    xlaCPP.Range("F" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 8).Value = r!Fechac
'    xlaCPP.Range("H" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'
'    xlaCPP.Cells(j, 9).Value = r!fechaini
'    xlaCPP.Range("I" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'
'    j = j + 1
'    i = i + 1
'    cdtotal = cdtotal + r!credito
'    r.MoveNext
'    If i = 20 And Not r.EOF Then
'        a = "B"
'        b = "C"
'        k = 2
'        xlaCPP.Cells(j, 1).Value = "       Folio"
'        xlaCPP.Range("A" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 3).Value = "Nombre"
'        xlaCPP.Range("C" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 6).Value = "      Crédito"
'        xlaCPP.Range("F" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 8).Value = "      Alta"
'        xlaCPP.Range("H" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = False
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 9).Value = "     Inicio"
'        xlaCPP.Range("I" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        j = j + 1
'        i = 1
'    End If
'
'Wend
'r.Close
'
'xlaCPP.Cells(j + 1, 6).Value = "Total  Terminados"
''xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
'xlaCPP.Range("F" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
''xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
'xlaCPP.Range("H" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'j = j + 2
'
'xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
'    With xlaCPP.Selection.Borders(4)
'        .LineStyle = xlContinuous
'        .Weight = 3
'        .ColorIndex = xlAutomatic
'    End With
'
'j = j + 2
'
'''''''''''''''''''''''''''''''''''''''''''GASTOS
'
'xlaCPP.Cells(j, 1).Value = "Gastos"
''xlaCPP.Range("A" + CStr(j) + ":B" + CStr(j)).Select
'xlaCPP.Range("A" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 12
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 4).Value = "Fecha:"
''xlaCPP.Range("D" + CStr(j) + ":E" + CStr(j)).Select
'xlaCPP.Range("D" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 11
'xlaCPP.Selection.Font.Bold = True
'
'Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 4")
'If Not r.EOF Then
'    xlaCPP.Cells(j, 5).Value = IIf(IsNull(r!Fechag), " ", r!Fechag)
'    xlaCPP.Range("E" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = True
'    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
'End If
'
'j = j + 2
'
'xlaCPP.Cells(j, 2).Value = "Descripción"
'xlaCPP.Range("B" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j, 6).Value = "Importe"
'xlaCPP.Range("F" + CStr(j)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'j = j + 1
'cdtotal = 0
'While Not r.EOF
'
'    xlaCPP.Cells(j, 2).Value = r!descripcion
'    xlaCPP.Range("B" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    xlaCPP.Cells(j, 6).Value = Format(r!importe, "###,###,###,###0.00")
'    xlaCPP.Range("F" + CStr(j)).Select
'    xlaCPP.Selection.WrapText = False
'    xlaCPP.Selection.Orientation = 0
'    xlaCPP.Selection.AddIndent = False
'    xlaCPP.Selection.Font.Name = "Arial"
'    xlaCPP.Selection.Font.Size = 10
'    xlaCPP.Selection.Font.Bold = False
'    xlaCPP.Selection.MergeCells = True
'
'    j = j + 1
'    i = i + 1
'    cdtotal = cdtotal + r!importe
'    r.MoveNext
'    If i = 20 And Not r.EOF Then
'        a = "B"
'        b = "C"
'        k = 2
'        xlaCPP.Cells(j, 2).Value = "Descripción"
'        xlaCPP.Range("B" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        xlaCPP.Cells(j, 6).Value = "Importe"
'        xlaCPP.Range("F" + CStr(j)).Select
'        xlaCPP.Selection.WrapText = False
'        xlaCPP.Selection.Orientation = 0
'        xlaCPP.Selection.AddIndent = False
'        xlaCPP.Selection.MergeCells = True
'        xlaCPP.Selection.Font.Name = "Arial"
'        xlaCPP.Selection.Font.Size = 10
'        xlaCPP.Selection.Font.Bold = True
'
'        j = j + 1
'        i = 1
'    End If
'
'Wend
'r.Close
'
'xlaCPP.Cells(j + 1, 6).Value = "Total  Gastos"
''xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
'xlaCPP.Range("F" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
''xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
'xlaCPP.Range("H" + CStr(j + 1)).Select
'xlaCPP.Selection.WrapText = False
'xlaCPP.Selection.Orientation = 0
'xlaCPP.Selection.AddIndent = False
'xlaCPP.Selection.MergeCells = True
'xlaCPP.Selection.Font.Name = "Arial"
'xlaCPP.Selection.Font.Size = 10
'xlaCPP.Selection.Font.Bold = True
'
'j = j + 2
'
'xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
'    With xlaCPP.Selection.Borders(4)
'        .LineStyle = xlContinuous
'        .Weight = 3
'        .ColorIndex = xlAutomatic
'    End With
'
'j = j + 2
'bgral = 0
'archivo = "C:\General\General"
'archivo = archivo + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")" + ".xls"
'nombreglobal = archivo
'On Error GoTo fin1
'xlwCPP.Saveas (archivo)
'bgral = 1
'GoTo fin2
'fin1: MsgBox "El archivo generado no va a ser guardado", vbCritical, "Reporte General"
'
'fin2:
'End Sub

