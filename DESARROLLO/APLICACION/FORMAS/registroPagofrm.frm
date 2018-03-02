VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form registroPagofrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Pago"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegistro 
      Caption         =   "Registro Pago"
      Height          =   465
      Left            =   5100
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   6390
      TabIndex        =   9
      Top             =   5640
      Width           =   1215
   End
   Begin Threed.SSFrame fraPago 
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   810
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   1984
      _Version        =   196609
      ForeColor       =   -2147483635
      Caption         =   "Programa del "
      Begin EditLib.fpDoubleSingle flMonto 
         Height          =   345
         Left            =   210
         TabIndex        =   1
         Top             =   570
         Width           =   1605
         _Version        =   196608
         _ExtentX        =   2831
         _ExtentY        =   609
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
         ControlType     =   1
         Text            =   "0.00"
         DecimalPlaces   =   2
         DecimalPoint    =   "."
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
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
      Begin EditLib.fpDateTime dtFecha 
         Height          =   345
         Left            =   5970
         TabIndex        =   3
         Top             =   570
         Width           =   1575
         _Version        =   196608
         _ExtentX        =   2778
         _ExtentY        =   609
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
         AlignTextH      =   1
         AlignTextV      =   1
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   1
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
         ControlType     =   1
         Text            =   "31/10/2007"
         DateCalcMethod  =   0
         DateTimeFormat  =   0
         UserDefinedFormat=   ""
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label lblCuentaCheques 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1890
         TabIndex        =   2
         Top             =   570
         Width           =   4005
      End
      Begin VB.Label Label2 
         Caption         =   "Cuenta de cheques cargo"
         Height          =   255
         Left            =   1890
         TabIndex        =   11
         Top             =   330
         Width           =   2325
      End
      Begin VB.Label Label3 
         Caption         =   "Monto:"
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   330
         Width           =   1845
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha de pago:"
         Height          =   255
         Left            =   5970
         TabIndex        =   12
         Top             =   300
         Width           =   1335
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3525
      Left            =   0
      TabIndex        =   13
      Top             =   1980
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   6218
      _Version        =   196609
      ForeColor       =   -2147483635
      Caption         =   "Datos del pago a realizar:"
      Begin VB.TextBox txtFolio 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   2355
      End
      Begin FPSpread.vaSpread sprParciales 
         Height          =   1725
         Left            =   180
         TabIndex        =   6
         Top             =   1140
         Width           =   7365
         _Version        =   196608
         _ExtentX        =   12991
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
         MaxCols         =   5
         MaxRows         =   50
         SpreadDesigner  =   "registroPagofrm.frx":0000
         UserResize      =   1
      End
      Begin SSCalendarWidgets_A.SSDateCombo dtFechaPago 
         Height          =   345
         Left            =   5760
         TabIndex        =   5
         Top             =   600
         Width           =   1785
         _Version        =   65537
         _ExtentX        =   3149
         _ExtentY        =   609
         _StockProps     =   93
         MinDate         =   "2007/1/1"
         MaxDate         =   "2010/12/31"
         Mask            =   2
      End
      Begin EditLib.fpDoubleSingle fTotalPorPagar 
         Height          =   345
         Left            =   5310
         TabIndex        =   7
         Top             =   3090
         Width           =   2235
         _Version        =   196608
         _ExtentX        =   3942
         _ExtentY        =   609
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
         ControlType     =   1
         Text            =   "0.00"
         DecimalPlaces   =   2
         DecimalPoint    =   "."
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
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
      Begin VB.Label Label10 
         Caption         =   "Folio:"
         Height          =   225
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label6 
         Caption         =   "Total a pagar:"
         Height          =   195
         Left            =   4290
         TabIndex        =   15
         Top             =   3090
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha de pago:"
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   360
         Width           =   1725
      End
   End
End
Attribute VB_Name = "registroPagofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COL_CUENTA_CHEQUES = 1
Const COL_SALDO = 2
Const COL_FORMA_PAGO = 3
Const COL_CHEQUE_TRANSFERENCIA = 4
Const COL_MONTO = 5

Public strConcepto As String
Public iCuentaContable As Integer
Public iCuentaBanco As Integer
Public iPago As Integer
Public iPagoConsecutivo As Integer
Public dMonto As Double
Public dMontoDelPago As Double
Public strFecha As String
Public iConcepto As Integer

Dim striCuenta As String
Dim strCuenta As String

Private Sub Form_Load()
    
    Dim oCtaCheques As New CuentaCheques
    
    If oCtaCheques.catalogoEsp() = True Then
        Call llenaComboSpread(sprParciales, COL_CUENTA_CHEQUES, oCtaCheques.cDatos, 1)
    End If
    
    fraPago.Caption = fraPago.Caption + strConcepto
    
    flMonto.Text = Format(dMonto, "$#,###.00")
    dtFecha.Text = strFecha
        
    dtFechaPago.Text = Date
    
    lblCuentaCheques.Caption = oCtaCheques.nombreCuenta(iCuentaBanco)
    
    Set oCtaCheques = Nothing
    
    Dim oPago As New cProveedor
    If oPago.formas() Then
        Call llenaComboSpread(sprParciales, COL_FORMA_PAGO, oPago.cDatos, 1)
    End If
    Set oPago = Nothing
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub sprParciales_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

    sprParciales.Col = Col
    sprParciales.Row = Row
    sprParciales.TypeComboBoxIndex = sprParciales.TypeComboBoxCurSel
        
    Dim iPos As Integer
    Dim dSaldoCuenta As Double
    Dim oCuentaCheques As New CuentaCheques
    
    Select Case Col
        Case Is = COL_CUENTA_CHEQUES
            
            striCuenta = Mid(sprParciales.TypeComboBoxString, 1, 2)
            strCuenta = Mid(sprParciales.TypeComboBoxString, InStr(1, sprParciales.TypeComboBoxString, " ") + 1)
            
            Call oCuentaCheques.saldoCuenta(Val(striCuenta), dSaldoCuenta)
            sprParciales.Col = COL_SALDO
            sprParciales.Text = dSaldoCuenta
            
        Case Is = COL_FORMA_PAGO
            
            If striCuenta = "" Then
                MsgBox "Seleccione una cuenta de cheques !!! ", vbInformation + vbOKOnly
                Exit Sub
            Else
                Dim strModo As String
                
                iPos = InStr(sprParciales.TypeComboBoxString, "-")
                strModo = Mid(sprParciales.TypeComboBoxString, iPos + 1)
                
                sprParciales.Col = COL_CHEQUE_TRANSFERENCIA
                If strModo = "Cheque" Then
                    Dim iChequeDisponible As Long
                    
                    If oCuentaCheques.siguienteDisponible(Val(striCuenta), iChequeDisponible) = True Then
                        sprParciales.Text = iChequeDisponible
                    Else
                        If vbYes = MsgBox("De la cuenta seleccionada, no hay cheques disponibles." & Chr(13) & "Para registrar cheques seleccione Yes" & Chr(13) & "Si desea por ahora hacer el pago por transferencia, seleccione No", vbInformation + vbYesNo) Then
                            
                            mantenimientoChequesfrm.iCuentaCheques = Val(striCuenta)
                            mantenimientoChequesfrm.strCuentaCheques = strCuenta
                            mantenimientoChequesfrm.Show vbModal
                            If oCuentaCheques.siguienteDisponible(Val(striCuenta), iChequeDisponible) = True Then
                                sprParciales.Col = COL_CHEQUE_TRANSFERENCIA
                                sprParciales.Text = iChequeDisponible
                            End If
                        End If
                    End If
                    
                Else
                    sprParciales.Text = ""
                End If
            End If
            
    End Select
    
    Set oCuentaCheques = Nothing
    
End Sub

Private Sub sprParciales_KeyPress(KeyAscii As Integer)

    If vbKeyReturn = KeyAscii Then
    
        Select Case sprParciales.ActiveCol
            
            Case Is = COL_MONTO
                Dim dMonto As Double
                Dim dSaldo As Double
                Dim dPago As Double
                Dim dTotalPorPagar As Double
                
                sprParciales.Col = COL_SALDO
                sprParciales.Row = sprParciales.ActiveRow
                'sprParciales.Action = ActionActiveCell
                dSaldo = Val(fnstrValor(sprParciales.Text))
                
                sprParciales.Col = COL_MONTO
                'sprParciales.Action = ActionActiveCell
                dMonto = Val(fnstrValor(sprParciales.Text))
                
                dPago = Val(fnstrValor(flMonto.Text))
                
                If dMonto > dSaldo Then
                    MsgBox "SALDO INSUFICIENTE EN LA CUENTA, VERIFIQUE POR FAVOR!!! ", vbCritical + vbOKOnly
                    Exit Sub
                End If
                
                If dMonto > dSaldo And dMonto > dPago Then
                    MsgBox "VERFIQUE, EL MONTO A LIQUIDAR ES MAYOR AL PAGO POR REALIZAR!!! ", vbCritical + vbOKOnly
                    Exit Sub
                End If
                
                If dMonto > dPago Then
                    MsgBox "EL VALOR DEL CHEQUE/TRANSFERENCIA, NO PUEDE SER MAYOR QUE EL PAGO POR REALIZAR, VERIFIQUE POR FAVOR", vbCritical + vbOKOnly
                    Exit Sub
                End If
                
                fTotalPorPagar = obtenTotalGrid(sprParciales, COL_MONTO)
                dTotalPorPagar = Val(fnstrValor(fTotalPorPagar))
                
                If dTotalPorPagar > dPago Then
                    MsgBox "VERFIQUE, EL MONTO A LIQUIDAR ES MAYOR AL PAGO POR REALIZAR!!! ", vbCritical + vbOKOnly
                    Exit Sub
                End If
                
        End Select
        
    End If
    
End Sub

Private Sub cmdRegistro_Click()

    Dim dPago As Double
    Dim dTotalPorPagar As Double
    
    dTotalPorPagar = Val(fnstrValor(fTotalPorPagar))
    dPago = Val(fnstrValor(flMonto))
    
    If txtFolio.Text = "" Then
        If MsgBox("NO HA DEFINIDO FOLIO (FACTURA O REMISION)" & Chr(13) & "Si desea registrar pago sin folio, seleccione Yes" & Chr(13) & "Si desea registrar el pago con un Folio, seleccione No", vbInformation + vbYesNo) = vbNo Then
            txtFolio.SetFocus
            Exit Sub
        End If
    End If
    If dTotalPorPagar > dPago Then
        If MsgBox("VERIFIQUE, EL MONTO DEL ADEUDO ES MAYOR AL MONTO DEL PAGO POR REALIZAR." & Chr(13) & "¿Aun desea realizar el pago?", vbQuestion + vbYesNo) = vbYes Then
            
            'REGISTRA EL PAGO
            Call registraPago
            Unload Me
            
        Else
            Exit Sub
        End If
    Else
    
        If dTotalPorPagar < dPago Then
            If MsgBox("El monto a liquidar es menor al pago. ¿Desea registrar el resto como un pago por pagar?", vbQuestion + vbYesNo) = vbYes Then
                'REGISTRAR EL PAGO
                Call registraPago
                
                'REGISTRAR EL SALDO COMO UN PAGO POR PAGAR
                pagoProrrogafrm.lblConcepto.Caption = strConcepto
                pagoProrrogafrm.fTotalPorPagar.Text = dPago - dTotalPorPagar
                pagoProrrogafrm.iPago = iPago
                pagoProrrogafrm.Show vbModal
                
                Unload Me
                
            Else
                'REGISTRAR EL PAGO
                Call registraPago
                Unload Me
                
            End If
        Else
            'REGISTRAR EL PAGO
            Call registraPago
            Unload Me
            
        End If
        
    End If
    
End Sub

Private Function registraPago()

    Dim lRow As Long
    Dim iCuentaBanco As Integer
    Dim iFormaPago As Integer
    Dim dMontoCheque As Double
    Dim strCheque As String
    
    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    Dim cPagos As New Collection
    Dim cPago As New Collection
    'cPago.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
    cPago.Add oCampo.CreaCampo(adInteger, , , iPago)
    cPago.Add oCampo.CreaCampo(adInteger, , , iPagoConsecutivo)
    cPago.Add oCampo.CreaCampo(adInteger, , , dtFecha.Text)
    cPago.Add oCampo.CreaCampo(adInteger, , , fnstrValor(fTotalPorPagar))
    cPago.Add oCampo.CreaCampo(adInteger, , , dtFechaPago.Text)
    cPago.Add oCampo.CreaCampo(adInteger, , , txtFolio.Text)
    cPagos.Add cPago
    
    For lRow = 1 To sprParciales.DataRowCnt
    
        Set cCheque = New Collection
        
        sprParciales.Row = lRow
        
        sprParciales.Col = COL_CUENTA_CHEQUES
        iCuentaBanco = Val(Left(sprParciales.Text, 2))
        
        sprParciales.Col = COL_FORMA_PAGO
        iFormaPago = Val(Left(sprParciales.Text, 2))
        
        sprParciales.Col = COL_MONTO
        dMontoCheque = fnstrValor(sprParciales.Text)
        
        sprParciales.Col = COL_CHEQUE_TRANSFERENCIA
        strCheque = sprParciales.Text
        
        'cCheque.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
        cCheque.Add oCampo.CreaCampo(adInteger, , , iCuentaBanco)
        cCheque.Add oCampo.CreaCampo(adInteger, , , dtFechaPago.Text)
        cCheque.Add oCampo.CreaCampo(adInteger, , , dMontoCheque)
        cCheque.Add oCampo.CreaCampo(adInteger, , , strCheque)
        cCheque.Add oCampo.CreaCampo(adInteger, , , iFormaPago)
        cCheque.Add oCampo.CreaCampo(adInteger, , , strConcepto)
        cCheque.Add oCampo.CreaCampo(adInteger, , , txtFolio.Text) 'factura del documento del gasto
        
        cCheques.Add cCheque
        
    Next lRow
    
    Dim oProveedor As New cProveedor
    'Dim strMensaje As String
    Call oProveedor.pagoRegistra(cPagos, cCheques, iConcepto, strConcepto)
    'strMensaje = oProveedor.pagoRegistra(gAlmacen, cPagos, cCheques, iConcepto, strConcepto)
    'Select Case strMensaje
    '    Case Is = "NO_EXISTE_CUENTA_CONTABLE_CHEQUES"
    '        MsgBox "¡La cuenta de cheques no esta relacionada a una cuenta contable!" + Chr(13) + "Su pago NO se registró por esta razón.", vbInformation + vbOKOnly
    '    Case Is = "NO_EXISTE_CUENTA_IVA_ACREDITABLE"
    '        MsgBox "¡No hay una cuenta contable de IVA acreditable, deber crear esta en el módulo contable!" + Chr(13) + "Su pago NO se registró por esta razón.", vbInformation + vbOKOnly
    '    Case Is = "NO_EXISTE_CUENTA_CONCEPTO"
    '        MsgBox "¡El concepto de pago, no esta relacionado a una cuenta contable!" + Chr(13) + "Su pago NO se registró por esta razón.", vbInformation + vbOKOnly
    '    Case Else
            MsgBox "¡Su pago quedó registrado!", vbInformation + vbOKOnly
    'End Select
    Set oProveedor = Nothing
    
'    'Asiento contable de la cuenta de cheques
'
'    'verifica si esta activo el modulo contable
'    If oAlmacen.contabilidadAbilitada = "SI" Then
'        Dim iCuentaContableCheques As Integer
'        Dim oCuentaCheques As New CuentaCheques
'        Call oCuentaCheques.cuentaContable(gAlmacen, iCuentaBanco, iCuentaContableCheques)
'        Set oCuentaCheques = Nothing
'
'        If iCuentaContableCheques = 0 Then
'            'Debe permitir dar de alta la cuenta contable y asociar esta a la cuenta de cheques.
'            If MsgBox("¡La chequera no tiene cuenta contable asociada o no existe!" + Chr(13) + "Si desea dar de alta esta o asociar a una, seleccione Yes" + Chr(13) + "Pero si prefiere hacerlo después, seleccione No" + Chr(13) + "ADVERTENCIA:Si NO da alta la cuenta contable, no se registrará el asiento contable.", vbInformation + vbYesNo) = vbYes Then
'                frmCuentas.bAlta = True
'                frmCuentas.strConceptoNuevo = lblCuentaCheques.Caption
'                frmCuentas.Show vbModal
'                iCuentaContableCheques = frmCuentas.iCuentaContable
'
'                Dim cRelacion As Collection
'                Dim cRelaciones As New Collection
'
'                'Asiento contable de la cuenta de cheques
'                Set cRelacion = New Collection
'                cRelacion.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
'                cRelacion.Add oCampo.CreaCampo(adInteger, , , iCuentaBanco) 'id de la cuenta contable
'                cRelacion.Add oCampo.CreaCampo(adInteger, , , iCuentaContableCheques)
'                cRelaciones.Add cRelacion
'
'                'Registra la relación de la cuenta contable y la cuenta de cheques
'                Call oCuentaCheques.registraRelacion(cRelaciones)
'
'            Else
'                Exit Function
'            End If
'
'        End If
'
'        If iCuentaContableCheques > 0 Then
'
'            Dim cAsiento As Collection
'            Dim cAsientos As New Collection
'
'            'Asiento contable de la cuenta de cheques
'            Set cAsiento = New Collection
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , iCuentaContableCheques) 'id de la cuenta contable
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , dtFechaPago.Text)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , dMontoCheque) 'Cargo
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , 0) 'Abono
'            cAsientos.Add cAsiento
'
'            'Asiento contable de la cuenta de gastos
'            Set cAsiento = New Collection
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , iCuentaContable) 'id de la cuenta contable
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , dtFechaPago.Text)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , 0) 'Cargo
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , dMontoCheque) 'Abono
'            cAsientos.Add cAsiento
'
'            'Asiento contable del IVA acreditable
'            'Set cAsiento = New Collection
'            'cAsiento.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
'            'cAsiento.Add oCampo.CreaCampo(adInteger, , , iCuentaContableIvaAcreditable) 'id de la cuenta contable
'            'cAsiento.Add oCampo.CreaCampo(adInteger, , , dtFechaPago.Text)
'            'cAsiento.Add oCampo.CreaCampo(adInteger, , , dMontoIVA) 'Cargo
'            'cAsiento.Add oCampo.CreaCampo(adInteger, , , 0) 'Abono
'            'cAsientos.Add cAsiento
'
'            'Realiza asiento contable
'            Dim oCuenta As New Cuenta
'            Call oCuenta.registraAsiento(cAsientos)
'            Set oCuenta = Nothing
'
'        End If
'
'    End If
    
End Function
