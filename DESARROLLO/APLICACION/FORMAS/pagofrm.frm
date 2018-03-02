VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form pagofrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago Recurrente"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Regitro Nuevo"
      Height          =   465
      Left            =   5550
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2145
      Left            =   0
      TabIndex        =   15
      Top             =   1980
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   3784
      _Version        =   196609
      ForeColor       =   -2147483635
      Caption         =   "Proyección de pagos"
      Begin VB.CheckBox chSiDias 
         Caption         =   "Deseo que en la lista de pagos aparezca antes que la fecha de pago el siguiente número de dias;"
         Height          =   735
         Left            =   3390
         TabIndex        =   9
         Top             =   1260
         Width           =   3675
      End
      Begin VB.ComboBox cbPeriodicidad 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   3195
      End
      Begin EditLib.fpLongInteger itxtNumeroPagosResta 
         Height          =   375
         Left            =   2340
         TabIndex        =   8
         Top             =   1410
         Width           =   615
         _Version        =   196608
         _ExtentX        =   1085
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         ButtonStyle     =   1
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
         AutoBeep        =   -1  'True
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
         Text            =   "0"
         MaxValue        =   "999"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
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
      Begin SSCalendarWidgets_A.SSDateCombo dtProximoPago 
         Height          =   345
         Left            =   210
         TabIndex        =   5
         Top             =   570
         Width           =   1815
         _Version        =   65537
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   93
         MinDate         =   "2007/1/1"
         MaxDate         =   "2010/12/31"
         Mask            =   2
      End
      Begin SSCalendarWidgets_A.SSDateCombo dtUltimoPago 
         Height          =   345
         Left            =   2520
         TabIndex        =   6
         Top             =   570
         Width           =   1815
         _Version        =   65537
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   93
         MinDate         =   "2007/1/1"
         MaxDate         =   "2010/12/31"
         Mask            =   2
      End
      Begin EditLib.fpLongInteger itxtDiasAntes 
         Height          =   375
         Left            =   7380
         TabIndex        =   10
         Top             =   1440
         Width           =   615
         _Version        =   196608
         _ExtentX        =   1085
         _ExtentY        =   661
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         ButtonStyle     =   1
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
         AutoBeep        =   -1  'True
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
         Text            =   "0"
         MaxValue        =   "999"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
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
      Begin VB.Label Label8 
         Caption         =   "Periodicidad"
         Height          =   165
         Left            =   4800
         TabIndex        =   19
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha de último pago:"
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de Próximo pago:"
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label6 
         Caption         =   "Número de pagos restantes:"
         Height          =   255
         Left            =   210
         TabIndex        =   16
         Top             =   1470
         Width           =   2115
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   6840
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   3413
      _Version        =   196609
      ForeColor       =   -2147483635
      Caption         =   "Definición de pago recurrente:"
      Begin VB.TextBox txtFolio 
         Height          =   375
         Left            =   2670
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ComboBox cbConcepto 
         Height          =   315
         Left            =   3930
         TabIndex        =   1
         Top             =   570
         Width           =   3405
      End
      Begin VB.ComboBox cbFormaPago 
         Height          =   315
         Left            =   5310
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1290
         Width           =   2745
      End
      Begin EditLib.fpDoubleSingle dlMonto 
         Height          =   315
         Left            =   3930
         TabIndex        =   3
         Top             =   1290
         Width           =   1065
         _Version        =   196608
         _ExtentX        =   1879
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Text            =   "0.00"
         DecimalPlaces   =   2
         DecimalPoint    =   "."
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
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
      Begin VB.ComboBox cbCuentaCheques 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1290
         Width           =   3435
      End
      Begin Threed.SSCommand cmdProveedor 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   23
         ToolTipText     =   "Captura del Teléfono..."
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
         CaptionStyle    =   1
         MarqueeDirection=   2
         BackStyle       =   1
         Caption         =   "Proveedor ..."
         Alignment       =   1
         ButtonStyle     =   3
         Outline         =   0   'False
      End
      Begin VB.Label lblRazonSocial 
         Caption         =   "Proveedor"
         Height          =   225
         Left            =   240
         TabIndex        =   22
         Top             =   630
         Width           =   3645
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto:"
         Height          =   225
         Left            =   3930
         TabIndex        =   21
         Top             =   300
         Width           =   2325
      End
      Begin VB.Label Label9 
         Caption         =   "Forma de Pago:"
         Height          =   255
         Left            =   5310
         TabIndex        =   20
         Top             =   990
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "Monto:"
         Height          =   195
         Left            =   3930
         TabIndex        =   14
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta de cheques cargo"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1020
         Width           =   2325
      End
   End
End
Attribute VB_Name = "pagofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bCambio As Boolean
Private iProveedor As Integer

Private Sub Form_Load()
   
    'Carga catálogo CUENTAS DE CHEQUES
    Dim oCtaCheques As New CuentaCheques
    
    If oCtaCheques.catalogoEsp = True Then
        Call fnLlenaComboCollecion(cbCuentaCheques, oCtaCheques.cDatos, 0, "")
    Else
        
        MsgBox "¡Es importante registrar una cuenta de cheques, para el control de sus finanzas!", vbInformation + vbOKOnly
        
        cuentaChequesfrm.Show vbModal
        
        If oCtaCheques.catalogoEsp = True Then
            Call fnLlenaComboCollecion(cbCuentaCheques, oCtaCheques.cDatos, 0, "")
        Else
            MsgBox "¡Para el control de sus finanzas, registre por lo menos una cuenta de cheques!", vbCritical + vbOKOnly
            Exit Sub
        End If
        
    End If
    
    Set oCtaCheques = Nothing
   
    'Carga catálogo GASTOS
    Dim oProveedor As New cProveedor
    
    If oProveedor.catalogoGastos(ID_OPERACION_GASTOS_OPERACION) Then
        Call fnLlenaComboCollecion(cbConcepto, oProveedor.cDatos, 0, "")
    Else
        
        oProveedor.creaModulo (gAlmacen)
        
        If oProveedor.catalogoGastos(ID_OPERACION_GASTOS_OPERACION) Then
            Call fnLlenaComboCollecion(cbConcepto, oProveedor.cDatos, 0, "")
        End If
    End If
    Set oProveedor = Nothing

    'Carga catálogo de periodos de pago
    Dim oPago As New cProveedor
    If oPago.periodicidad() Then
        Call fnLlenaComboCollecion(cbPeriodicidad, oPago.cDatos, 0, "")
    End If

    'Carga catálogo de formas de pago
    If oPago.formas() Then
        Call fnLlenaComboCollecion(cbFormaPago, oPago.cDatos, 0, "")
    End If
    Set oPago = Nothing

End Sub

'Private Sub cmdProveedor_Click(Index As Integer)
'
'    listaProveedoresfrm.Show vbModal
'    iProveedor = Val(listaProveedoresfrm.mColUno)
'    lblRazonSocial.Caption = Val(listaProveedoresfrm.mColDos)
'    Unload listaProveedoresfrm
'
'End Sub


Private Sub chSiDias_Click()
    itxtDiasAntes.Enabled = chSiDias.Value
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    
    'Dim iCuentaContable As Integer
    Dim oPago As New cProveedor
    
    If cbCuentaCheques.ListIndex = -1 Then
        MsgBox "Defina por favor la cuenta de cheques (Cargo)", vbInformation + vbOKOnly
        cbCuentaCheques.SetFocus
        Exit Sub
    End If
    
    If cbFormaPago.ListIndex = -1 Then
        MsgBox "Defina por favor la Forma de Pago", vbInformation + vbOKOnly
        cbFormaPago.SetFocus
        Exit Sub
    End If
    
    If cbPeriodicidad.ListIndex = -1 Then
        MsgBox "Defina por favor Periodicidad del pago.", vbInformation + vbOKOnly
        cbPeriodicidad.SetFocus
        Exit Sub
    End If
    
    If cbConcepto.ListIndex = -1 Then  ' pago nuevo

        If oPago.buscaConNombre(gAlmacen, cbConcepto.Text) = True Then
            
            MsgBox "El pago ya existe, si desea cambiar condiciones seleccione la opción Editar del modulo Pagos", vbCritical + vbInformation
            
        'Else
            'frmCuentas.bAlta = True
            'frmCuentas.strConceptoNuevo = cbConcepto.Text
            'frmCuentas.Show vbModal
            'iCuentaContable = frmCuentas.iCuentaContable
        End If
    Else
        Dim iConcepto As Integer
        iConcepto = cbConcepto.ItemData(cbConcepto.ListIndex)
    End If
            
    'Verifica si hay modulo contable
    'Dim strRespuesta As String
    'Dim iModuloContable As Integer
    'iModuloContable = 0
    
    'strRespuesta = oAlmacen.contabilidadAbilitada()
    'If "SI" = strRespuesta Then
    '    iModuloContable = 1
    'ElseIf "SI_PRIMER_VEZ" = strRespuesta Then
        'Implica crear catalogo y actualizar relación tipo operació vs cuenta
    '    Dim oCuenta As New Cuenta
    '    oCuenta.actualizaRelaciones (gAlmacen)
    '    Set oCuenta = Nothing
    'End If
    
    Dim cPago As New Collection
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    Dim iPos As Integer
    Dim iBanco As Integer
    
    iPos = InStr(1, cbCuentaCheques.Text, " ")
    iBanco = Val(Mid(cbCuentaCheques.Text, 1, iPos))
    
    Set cRegistro = New Collection
    'cRegistro.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , iConcepto)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , cbCuentaCheques.ItemData(cbCuentaCheques.ListIndex))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , cbConcepto.Text)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , dtProximoPago.Text)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , dtUltimoPago.Text)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , itxtDiasAntes.Text)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , itxtNumeroPagosResta.Text)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , cbFormaPago.ItemData(cbFormaPago.ListIndex))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , cbPeriodicidad.ItemData(cbPeriodicidad.ListIndex))
    'cRegistro.Add oCampo.CreaCampo(adInteger, , , iModuloContable)
    cPago.Add cRegistro
    
    Call oPago.creaNuevo(cPago, _
                         dlMonto.Text, _
                         dtProximoPago.Text, _
                         dtUltimoPago.Text, _
                         cbPeriodicidad.ItemData(cbPeriodicidad.ListIndex), _
                         itxtNumeroPagosResta.Text, _
                         txtfolio.Text)
            
    Set oPago = Nothing
    
    Unload Me
                       
End Sub
