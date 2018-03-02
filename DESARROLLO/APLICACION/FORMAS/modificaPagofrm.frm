VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form modificaPagofrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modifica Pagos"
   ClientHeight    =   7515
   ClientLeft      =   1860
   ClientTop       =   735
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdgraba 
      Caption         =   "      Grabar Modificaciones"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   6990
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crédito"
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtcte 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   810
         TabIndex        =   12
         Top             =   1215
         Width           =   1410
      End
      Begin VB.TextBox txtfolio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   420
         Left            =   810
         TabIndex        =   11
         Top             =   315
         Width           =   1455
      End
      Begin EditLib.fpLongInteger txtfolioold 
         Height          =   435
         Left            =   810
         TabIndex        =   7
         Top             =   300
         Width           =   1425
         _Version        =   196608
         _ExtentX        =   2514
         _ExtentY        =   767
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         Text            =   "0"
         MaxValue        =   "2147483647"
         MinValue        =   "1"
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
      Begin VB.TextBox txtnombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   1200
         Width           =   3615
      End
      Begin EditLib.fpLongInteger txtcteold 
         Height          =   285
         Left            =   810
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1425
         _Version        =   196608
         _ExtentX        =   2514
         _ExtentY        =   503
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
         MaxValue        =   "2147483647"
         MinValue        =   "1"
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
      Begin VB.Label Label1 
         Caption         =   "Folio:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   400
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   975
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "No Cliente:"
         Height          =   255
         Left            =   810
         TabIndex        =   3
         Top             =   915
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pagos Registrados"
      Height          =   4845
      Left            =   0
      TabIndex        =   0
      Top             =   2070
      Width           =   6735
      Begin FPSpread.vaSpread sprPagos 
         Height          =   4425
         Left            =   270
         TabIndex        =   6
         Top             =   330
         Width           =   6405
         _Version        =   196608
         _ExtentX        =   11298
         _ExtentY        =   7805
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         SpreadDesigner  =   "modificaPagofrm.frx":0000
      End
   End
   Begin Threed.SSPanel pnlMensaje 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   1590
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   873
      _Version        =   196608
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
   End
End
Attribute VB_Name = "modificaPagofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_FECHA = 1
Private Const COL_PAGO = 2
Private Const COL_ADEUDO = 3
Private Const COL_HORA = 4
Private Const COL_USUARIO = 5
Private Const COL_LUGAR = 6
Private Const COL_NO_CLIENTE = 7
Private Const COL_NOMBRE_CLIENTE = 8
Private Const COL_ORDEN = 9
Private Const COL_CONSECUTIVO = 10
Private Const COL_MODIFICA = 11

Private lRowActiva As Long

Private fAdeudo As Double

Public lFolio As Long

Private fPagoActualTemp As Double

Private fPago As Double
Private bCambio As Boolean

Private Sub Form_Activate()

    txtfolio.Text = lFolio
    txtfolio_KeyPress (vbKeyReturn)
    
    If giTipoUsuario <> USUARIO_GERENTE Then
    
        cmdgraba.Enabled = False
        sprPagos.Col = COL_PAGO
        sprPagos.Col2 = COL_PAGO
        sprPagos.BlockMode = True
        sprPagos.Lock = True
        sprPagos.BlockMode = False
        
    Else
        cmdgraba.Enabled = True
    End If
    
End Sub

Private Sub txtfolio_GotFocus()
    pnlMensaje = ""
    txtfolio.SelStart = 0
    txtfolio.SelLength = Len(txtfolio.Text)
End Sub

Private Sub txtfolio_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        Dim oPago As New Pago
        
        If oPago.obtenCreditoPagos(Val(txtfolio)) Then
            
            Dim fAdeudo, fPago As Double
            Dim lRow As Long
            
            'Asegura que es negro todo
            sprPagos.Row = -1
            sprPagos.Col = -1
            sprPagos.ForeColor = RGB(0, 0, 0) '&H8000&
            
            pnlMensaje = ""
            Call fnLlenaTablaCollection(sprPagos, oPago.cDatos)
            
            sprPagos.Col = COL_ADEUDO
            sprPagos.Row = 1
            
            fAdeudo = Val(fnstrValor(sprPagos.Text))
            
            For lRow = 1 To sprPagos.DataRowCnt
                
                sprPagos.Col = COL_PAGO
                sprPagos.Row = lRow
                fPago = Val(fnstrValor(sprPagos.Text))
                
                sprPagos.Col = COL_ADEUDO
                fAdeudo = fAdeudo - fPago
                sprPagos.Text = fAdeudo
                
            Next lRow
            
            sprPagos.Col = COL_NO_CLIENTE
            sprPagos.Row = 1
            
            txtcte = sprPagos.Text
            sprPagos.Col = COL_NOMBRE_CLIENTE
            txtnombre = sprPagos.Text
            
            sprPagos.SetFocus
            
            sprPagos.Col = COL_FECHA
            sprPagos.Row = 1
            sprPagos.Action = ActionActiveCell
        Else
            
            pnlMensaje = "El folio que se esta buscando no existe"
            
            fnLimpiaGrid sprPagos
            txtcte.Text = ""
            txtnombre.Text = ""
            
            txtfolio.SetFocus
            
        End If
        
        Set oPago = Nothing
    
    End If
'    Dim datos As Recordset
'    Dim datos1 As Recordset
'    Dim bandera As Integer
'    Dim i As Integer
'    i = 1
'    If KeyAscii = 13 Then
'        If txtfolio.Text <> "" And IsNumeric(txtfolio.Text) Then
'            txtfolio.Text = CLng(txtfolio.Text)
'            grdpagos.Clear
'            grdpagos.Rows = 2
'            Call Form_Load
'            Set datos = base.OpenRecordset("select * from creditos where factura=" & CStr(txtfolio.Text))
'            If datos.RecordCount > 0 Then
'                txtcte.Text = datos!no_cliente
'                datos.Close
'                Set datos = base.OpenRecordset("select * from clientes where no_cliente=" & CStr(txtcte.Text))
'                If datos.RecordCount > 0 Then
'                    txtnombre.Text = datos!Nombre + " " + datos!apellido
'                    datos.Close
'                    Set datos1 = base.OpenRecordset("select * from pagos where no_cliente=" & CStr(txtcte.Text) & " and factura=" & CStr(txtfolio.Text))
'                    If datos1.RecordCount > 0 Then
'
'                        datos1.MoveFirst
'                        While Not datos1.EOF
'                            grdpagos.Row = grdpagos.Rows - 1
'                            grdpagos.Rows = grdpagos.Row + 2
'                            grdpagos.Col = 0
'                            grdpagos.Text = i
'                            grdpagos.Col = 1
'                            grdpagos.Text = datos1!fecha
'                            grdpagos.Col = 2
'                            grdpagos.Text = Format(datos1!Cantpagada, "###,###,###,###0.00")
'                            grdpagos.Col = 3
'                            grdpagos.Text = Format(datos1!Cantadeudada, "###,###,###,###0.00")
'                            grdpagos.Col = 4
'                            grdpagos.Text = datos1!hora
'                            grdpagos.Col = 5
'                            grdpagos.Text = datos1!usuario
'                            grdpagos.Col = 6
'                            grdpagos.Text = datos1!lugar
'                            datos1.MoveNext
'                            i = i + 1
'                        Wend
'                    End If
'                End If
'            Else
'                MsgBox "El folio que se esta buscando no existe", vbInformation, "Registro de Pagos"
'
'                txtcte.Text = ""
'                txtnombre.Text = ""
'                txtpago.Text = ""
'                txtadeudo.Text = ""
'                txtfecha.Text = Format(Now, "dd/mm/yyyy")
'                txtfolio.Text = ""
'                txtfolio.SetFocus
'
'            End If
'        End If
'
'    End If


End Sub

Private Sub sprPagos_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If giTipoUsuario = USUARIO_GERENTE Then
        sprPagos.Col = Col
        sprPagos.Row = Row
        
        If Mode = 1 Then
                       
            fPagoActualTemp = Val(fnstrValor(sprPagos.Text))
            
        Else
            
            fPago = Val(fnstrValor(sprPagos.Text))
            bCambio = True
        
        End If
    End If
    
End Sub

Private Sub sprPagos_GotFocus()
    
    If giTipoUsuario = USUARIO_GERENTE Then
    
        sprPagos.Row = sprPagos.ActiveRow
        sprPagos.Col = COL_PAGO
        
        fPagoActualTemp = Val(fnstrValor(sprPagos.Text))
    End If
    
End Sub

Private Sub sprPagos_Click(ByVal Col As Long, ByVal Row As Long)

    If giTipoUsuario = USUARIO_GERENTE Then

        sprPagos.Row = Row
        sprPagos.Col = COL_PAGO
        
        fPagoActualTemp = Val(fnstrValor(sprPagos.Text))
    End If
    
End Sub

Private Sub sprPagos_LostFocus()

    If giTipoUsuario = USUARIO_GERENTE Then
    
        sprPagos.Row = sprPagos.ActiveRow
        sprPagos.Col = COL_PAGO
            
        fPago = Val(fnstrValor(sprPagos.Text))
    End If
    
End Sub

Private Sub sprPagos_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    If giTipoUsuario = USUARIO_GERENTE Then

        If Row = NewRow Then
        
            Select Case Col
                
                Case Is = COL_FECHA
                
                    sprPagos.Row = Row
                    sprPagos.Col = COL_PAGO
                
                    fPagoActualTemp = Val(fnstrValor(sprPagos.Text))
        
                Case Is = COL_PAGO
                    
                     If bCambio = True Then
                        
                        sprPagos.Row = Row
                        sprPagos.Col = COL_PAGO
                        fPago = Val(fnstrValor(sprPagos.Text))
                    
                        If recalculaPagos = True Then
                           cmdgraba.Enabled = True
                        End If
                        
                        bCambio = False
                        
                     End If
                    
            End Select
        Else
        
            sprPagos.Row = Row
            sprPagos.Col = COL_PAGO
            fPago = Val(fnstrValor(sprPagos.Text)) 'El pago ya modificado
            
            If recalculaPagos = True Then
                cmdgraba.Enabled = True
                bCambio = False
            End If
            
            If NewRow <> -1 Then
                sprPagos.Row = NewRow
                sprPagos.Col = COL_PAGO
                fPagoActualTemp = Val(fnstrValor(sprPagos.Text))
            End If
            
        End If
    
    End If
    
End Sub

Private Function recalculaPagos() As Boolean

    recalculaPagos = False
    
    If fPago = fPagoActualTemp Then
        Exit Function
    End If
    
    Dim fAdeudo As Double
    Dim lRow As Long
    Dim fAdeudoAnterior As Double
    Dim fPagoActual As Double
    
    sprPagos.Col = COL_ADEUDO
    fAdeudo = Val(fnstrValor(sprPagos.Text))
    If fPago > fPagoActualTemp Then
        
        fAdeudo = fAdeudo - (fPago - fPagoActualTemp)
        
    Else
    
        fAdeudo = fAdeudo + (fPagoActualTemp - fPago)
        
    End If
    sprPagos.Text = fAdeudo
            
    'Indica que ha cambiado el pago de este
    sprPagos.Col = COL_MODIFICA
    sprPagos.Text = 1
    
    'Camba el color del texto
    sprPagos.Col = -1
    sprPagos.ForeColor = RGB(255, 0, 0) '&H8000&
            
    'Recalcula saldos
    For lRow = sprPagos.ActiveRow + 1 To sprPagos.DataRowCnt
        
        sprPagos.Row = lRow - 1
        sprPagos.Col = COL_ADEUDO
        fAdeudoAnterior = Val(fnstrValor(sprPagos.Text))
        
        sprPagos.Row = lRow
        sprPagos.Col = COL_PAGO
        fPagoActual = Val(fnstrValor(sprPagos.Text))
        
        sprPagos.Col = COL_ADEUDO
        sprPagos.Text = fAdeudoAnterior - fPagoActual
        
        sprPagos.Col = COL_MODIFICA
        sprPagos.Text = 1
        
    Next lRow
        
    recalculaPagos = True
End Function

Private Sub sprPagos_KeyPress(KeyAscii As Integer)

    If giTipoUsuario = USUARIO_GERENTE Then

        If KeyAscii = vbKeyReturn Then
            
            Dim lRowAnterior As Long
            
            Select Case sprPagos.ActiveCol
                
                Case Is = COL_PAGO
                                
                    sprPagos.Row = sprPagos.ActiveRow
                    lRowAnterior = sprPagos.ActiveRow
                    
                    fPago = Val(fnstrValor(sprPagos.Text))
                                
                    If recalculaPagos = True Then
                                    
                        cmdgraba.Enabled = True
                        bCambio = False
                    
                        sprPagos.Row = lRowAnterior
                        sprPagos.Col = COL_PAGO
                        fPagoActualTemp = Val(fnstrValor(sprPagos.Text))
    
                    End If
                    
            End Select
            
        End If
    
    End If
    
End Sub

Private Function obtenPagos() As Collection

    Dim lRenglon As Long
    Dim cPagos As New Collection
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    Dim bPagos As Boolean
    
    bPagos = False
    
    For lRenglon = 1 To sprPagos.DataRowCnt
    
        Set cRegistro = New Collection
        
        sprPagos.Row = lRenglon
        
        sprPagos.Col = COL_MODIFICA
        If 1 = Val(sprPagos.Text) Then
        
            cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtfolio.Text))
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , 0)
            sprPagos.Col = COL_PAGO
            cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPagos.Text)))
            sprPagos.Col = COL_FECHA
            cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPagos.Text)
            'sprPagos.Col = COL_USUARIO
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPagos.Text)
            'sprPagos.Col = COL_HORA
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPagos.Text)
            'sprPagos.Col = COL_LUGAR
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPagos.Text)
            sprPagos.Col = COL_CONSECUTIVO
            'sprPagos.Col = COL_ORDEN
            cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(sprPagos.Text))
            
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'INDICA QUE DEBE GRABAR EN PAGOS
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtcte.Text))
            sprPagos.Col = COL_ADEUDO
            cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPagos.Text)))
    
            cPagos.Add cRegistro
            bPagos = True
            
        End If
        
    Next lRenglon
    
    Set obtenPagos = cPagos
    
End Function

Private Sub cmdgraba_Click()

    Dim cPagos As New Collection
    
    Set cPagos = obtenPagos()
    
    If cPagos.Count > 0 Then
    
        Dim oPago As New Pago
        
        oPago.actualizaPagos cPagos
        
        Set oPago = Nothing
        
        MsgBox "La(s) modificacion(es) ha(n) sido realizada(s)", vbOKOnly, "Modificación de pagos"
        
        'Asegura que es negro todo
        sprPagos.Row = sprPagos.ActiveRow
        sprPagos.Col = -1
        sprPagos.ForeColor = RGB(0, 0, 0) '&H8000&
        
    End If
    
    'cmdmodifica.Enabled = False
    cmdgraba.Enabled = False
    
'    Dim datos As Recordset
'    Dim datos1 As Recordset
'    Dim j As Integer
'    Dim nocliente, folio As Long
'    Dim fecha As Date
'    Dim apago, aadeudo As Double
'    Dim hora, usuario, lugar As String
'
'    cmdmodifica.Enabled = False
'
'    base.Execute "delete from pagos where no_cliente=" & txtcte.Text & " and factura=" & txtfolio.Text
'
'    If txtfolio.Text <> "" And IsNumeric(txtfolio.Text) Then
'        folio = txtfolio.Text
'        nocliente = txtcte.Text
'
'        For j = 1 To grdpagos.Rows - 2
'            grdpagos.Row = j
'
'            grdpagos.Col = 1
'            fecha = CDate(grdpagos.Text)
'            grdpagos.Col = 2
'            apago = CDbl(grdpagos.Text)
'            grdpagos.Col = 3
'            aadeudo = CDbl(grdpagos.Text)
'            grdpagos.Col = 4
'            hora = grdpagos.Text
'            grdpagos.Col = 5
'            usuario = grdpagos.Text
'            grdpagos.Col = 6
'            lugar = grdpagos.Text
'
'
'            base.Execute "insert into pagos (no_cliente,factura,fecha,Cantpagada,Cantadeudada,cons_pago,usuario,hora,lugar) values(" + CStr(nocliente) + "," + CStr(folio) + ",'" + CStr(fecha) + "'," + CStr(apago) + "," + CStr(aadeudo) + ",1,'" + usuario + "','" + hora + "','" + lugar + "')"
'
'        Next j
'
'    End If
'
'    MsgBox "La(s) modificacion(es) ha(n) sido grabada(s)", vbOKOnly, "Modificación de pagos"
'    cmdmodifica.Enabled = False
'    cmdgraba.Enabled = False

End Sub

'Private Sub cmdSalir_Click()
'Unload Me
'End Sub

'Private Sub txtfolio_LostFocus()
'Call txtfolio_KeyPress(13)
'End Sub

'Private Sub grdpagos_DblClick()
'
'    renglon = grdpagos.Row
'    grdpagos.Col = 1
'    If grdpagos.Text <> "" Then
'        txtfecha.Text = grdpagos.Text
'        grdpagos.Col = 2
'        txtpago.Text = Format(CDbl(grdpagos.Text), "###,###,###,###0.00")
'        pago = grdpagos.Text
'        grdpagos.Col = 3
'        txtadeudo.Text = Format(CDbl(grdpagos.Text) + pago, "###,###,###,###0.00")
'        adeudo = grdpagos.Text
'
'        cmdmodifica.Enabled = True
'        cmdgraba.Enabled = False
'
'        txtpago.SetFocus
'    End If
'
'End Sub

'Private Sub cmdmodifica_Click()
'
'    Dim pagos, adeudos As Double
'
'    grdpagos.Row = renglon
'
'    If txtpago.Text <> "" And IsNumeric(txtpago.Text) Then
'        grdpagos.Col = 1
'        grdpagos.Text = txtfecha.Text
'        grdpagos.Col = 2
'        grdpagos.Text = Format(txtpago.Text, "###,###,###,###0.00")
'        grdpagos.Col = 3
'        grdpagos.Text = Format(CDbl(txtadeudo.Text) - CDbl(txtpago.Text), "###,###,###,###0.00")
'        txtadeudo.Text = Format(CDbl(txtadeudo.Text) - CDbl(txtpago.Text), "###,###,###,###0.00")
'        For i = renglon + 1 To grdpagos.Rows - 2
'            grdpagos.Row = i - 1
'            grdpagos.Col = 3
'            adeudos = CDbl(grdpagos.Text)
'
'            grdpagos.Row = i
'            grdpagos.Col = 2
'            pagos = CDbl(grdpagos.Text)
'            grdpagos.Col = 3
'            grdpagos.Text = Format(adeudos - pagos, "###,###,###,###0.00")
'        Next i
'
'        cmdgraba.Enabled = True
'        cmdmodifica.Enabled = False
'    Else
'        MsgBox "El dato del pago es incorrecto", vbCritical, "Registro de Pagos"
' '       txtpago.Text = pagogen
'        txtpago.SetFocus
'    End If
'End Sub

'Private Sub Form_Load()
'
'    grdpagos.Row = 0
'    grdpagos.Col = 1
'    grdpagos.Text = "Fecha"
'    grdpagos.Col = 2
'    grdpagos.Text = "Pago"
'    grdpagos.Col = 3
'    grdpagos.Text = "Adeudo"
'    grdpagos.Col = 4
'    grdpagos.Text = "Hora"
'    grdpagos.Col = 5
'    grdpagos.Text = "Usuario"
'    grdpagos.Col = 6
'    grdpagos.Text = "Lugar"
'
'End Sub

'Private Sub txtpago_GotFocus()
'    txtpago.SelStart = 0
'    txtpago.SelLength = Len(txtpago.Text)
'End Sub

'Private Sub txtpago_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Then
'        Dim fAdeudo, fPago As Double
'
'        sprPagos.Row = lRowActiva
'
'        sprPagos.Col = COL_ADEUDO
'        fAdeudo = Val(fnstrValor(sprPagos.Text))
'
'        sprPagos.Col = COL_PAGO
'        fPago = Val(fnstrValor(sprPagos.Text))
'
'        fAdeudo = fAdeudo + fPago 'l adeudo debe aumentar con el pago actual, para restar el nuevo pago y generar un nuevo adeudo
'
'        fPago = Val(fnstrValor(txtpago.Text)) 'El pago ya modificado
'
'        sprPagos.Col = COL_ADEUDO
'        sprPagos.Text = fAdeudo - fPago 'Resta al adeudo el nuevo pago
'
'        sprPagos.Col = COL_PAGO
'        sprPagos.Text = fPago   'Despliega el nuevo pago
'
'        cmdgraba.Enabled = True
'    End If
'
'End Sub

