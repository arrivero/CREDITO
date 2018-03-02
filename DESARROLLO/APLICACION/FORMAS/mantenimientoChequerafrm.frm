VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form mantenimientoChequesfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chequera"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlCheque 
      Height          =   4005
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   1200
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7064
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSFrame SSFrame2 
         Height          =   1005
         Left            =   300
         TabIndex        =   16
         Top             =   1530
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1773
         _Version        =   196609
         ForeColor       =   -2147483635
         Caption         =   "Registro Nuevo"
         Begin EditLib.fpLongInteger itxtChequeIncial 
            Height          =   405
            Left            =   450
            TabIndex        =   17
            Top             =   510
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2831
            _ExtentY        =   714
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
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
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger itxtChequeFinal 
            Height          =   405
            Left            =   2400
            TabIndex        =   18
            Top             =   510
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2831
            _ExtentY        =   714
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
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
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Inicial"
            Height          =   195
            Left            =   450
            TabIndex        =   20
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label3 
            Caption         =   "Final"
            Height          =   195
            Left            =   2400
            TabIndex        =   19
            Top             =   240
            Width           =   1635
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1065
         Left            =   300
         TabIndex        =   11
         Top             =   330
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1879
         _Version        =   196609
         ForeColor       =   -2147483635
         Caption         =   "Registro Actual"
         Begin EditLib.fpDoubleSingle itxtChequeAntInicial 
            Height          =   345
            Left            =   420
            TabIndex        =   14
            Top             =   570
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
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
            Text            =   "0"
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDoubleSingle itxtChequeAntFinal 
            Height          =   345
            Left            =   2370
            TabIndex        =   15
            Top             =   570
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
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
            Text            =   "0"
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label5 
            Caption         =   "Inicial"
            Height          =   195
            Left            =   420
            TabIndex        =   13
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label4 
            Caption         =   "Final"
            Height          =   195
            Left            =   2370
            TabIndex        =   12
            Top             =   240
            Width           =   1635
         End
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "Registrar"
         Height          =   525
         Left            =   4350
         TabIndex        =   3
         Top             =   2730
         Width           =   1215
      End
   End
   Begin Threed.SSPanel pnlCheque 
      Height          =   4005
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   1200
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7064
      _Version        =   196609
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.OptionButton rbCheque 
         Caption         =   "Cancelado(s)"
         Height          =   375
         Index           =   2
         Left            =   3570
         TabIndex        =   10
         Top             =   330
         Width           =   1515
      End
      Begin VB.OptionButton rbCheque 
         Caption         =   "Rebotado(s)"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   330
         Width           =   1515
      End
      Begin VB.OptionButton rbCheque 
         Caption         =   "Cobrado(s)"
         Height          =   375
         Index           =   0
         Left            =   750
         TabIndex        =   8
         Top             =   330
         Width           =   1515
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   525
         Left            =   4590
         TabIndex        =   7
         Top             =   3360
         Width           =   1215
      End
      Begin FPSpread.vaSpread sprCheques 
         Height          =   2355
         Left            =   210
         TabIndex        =   5
         Top             =   840
         Width           =   5625
         _Version        =   196608
         _ExtentX        =   9922
         _ExtentY        =   4154
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
         MaxCols         =   4
         MaxRows         =   100
         SpreadDesigner  =   "mantenimientoChequerafrm.frx":0000
         UserResize      =   1
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   525
      Left            =   4740
      TabIndex        =   6
      Top             =   5310
      Width           =   1215
   End
   Begin VB.ComboBox cbCuentaCheques 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   510
      Width           =   2775
   End
   Begin ComctlLib.TabStrip tabCheques 
      Height          =   4335
      Left            =   60
      TabIndex        =   22
      Top             =   900
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Registro"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Mantenimiento"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCuentaCheques 
      Caption         =   "Label6"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   540
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label Label2 
      Caption         =   "Cuenta de cheques:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   2325
   End
End
Attribute VB_Name = "mantenimientoChequesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miFrameActivo As Integer
Private iChequeActualIncial As Integer
Private iChequeActualfINAL As Integer
Private iEstadoNuevoCheque As Integer

Public iCuentaCheques As Integer
Public strCuentaCheques As String

Private Sub Form_Load()

    If iCuentaCheques = 0 Then
        'Carga catálogo CUENTAS DE CHEQUES
        Dim oCtaCheques As New CuentaCheques
        
        If oCtaCheques.catalogoEsp() = True Then
            Call fnLlenaComboCollecion(cbCuentaCheques, oCtaCheques.cDatos, 0, "")
        End If
        Set oCtaCheques = Nothing
    Else
        cbCuentaCheques.Visible = False
        lblCuentaCheques.Visible = True
        lblCuentaCheques.Caption = strCuentaCheques
        Call obtenRegistroActual(iCuentaCheques)
    End If
End Sub

Private Sub rbCheque_Click(Index As Integer)
    Select Case Index
        Case Is = 0
            
            If rbCheque.Item(Index).Value = True Then
                iEstadoNuevoCheque = ST_CHEQUE_COBRADO
            End If
            
        Case Is = 1
            If rbCheque.Item(Index).Value = True Then
                iEstadoNuevoCheque = ST_CHEQUE_REBOTADO
            End If
        Case Is = 2
            If rbCheque.Item(Index).Value = True Then
                iEstadoNuevoCheque = ST_CHEQUE_CANCELADO
            End If
    End Select
End Sub

Private Sub tabCheques_Click()
    
    If tabCheques.SelectedItem.Index - 1 = miFrameActivo Then Exit Sub ' No need to change frame.
    
    ' Comosea, oculta el frame anterior, muestra el nuevo.
    pnlCheque(tabCheques.SelectedItem.Index - 1).Visible = True
    pnlCheque(miFrameActivo).Visible = False
    
    miFrameActivo = tabCheques.SelectedItem.Index - 1

End Sub

Private Sub cbCuentaCheques_Click()

    iCuentaCheques = cbCuentaCheques.ItemData(cbCuentaCheques.ListIndex)
    Call obtenRegistroActual(iCuentaCheques)

    Call obtenEstatus(iCuentaCheques)

End Sub

'Private Function obtenEstatus(iSalon As Integer, iCuenta As Integer)
Private Function obtenEstatus(iCuenta As Integer)

    Dim oCuentaCheques As New CuentaCheques
    
    'Call oCuentaCheques.chequesEstatusActual(iSalon, _
    '                                         iCuenta, ST_CHEQUE_PAGADO)
    Call oCuentaCheques.chequesEstatusActual(iCuenta, ST_CHEQUE_PAGADO)
                                            
    Call fnLimpiaGrid(sprCheques)
    Call fnLlenaTablaCollection(sprCheques, oCuentaCheques.cDatos)
    
    Set oCuentaCheques = Nothing
    
End Function

'Private Function obtenRegistroActual(iSalon As Integer, iCuenta As Integer)
Private Function obtenRegistroActual(iCuenta As Integer)

    Dim oCuentaCheques As New CuentaCheques
    
'    Call oCuentaCheques.chequesRegistroActual(iSalon, _
'                                              iCuenta, _
'                                              iChequeActualIncial, _
'                                              iChequeActualfINAL)
    Call oCuentaCheques.chequesRegistroActual(iCuenta, _
                                              iChequeActualIncial, _
                                              iChequeActualfINAL)
                                            
    itxtChequeAntInicial.Text = iChequeActualIncial
    itxtChequeAntFinal.Text = iChequeActualfINAL
    
    Set oCuentaCheques = Nothing
    
End Function

Private Sub cmdRegistrar_Click()
    
    If Val(itxtChequeIncial.Text) <= 0 Then
        MsgBox "Defina cheque inicial", vbInformation + vbOKOnly
        itxtChequeIncial.SetFocus
        Exit Sub
    End If
    
    If Val(itxtChequeFinal.Text) <= 0 Then
        MsgBox "Defina cheque Final", vbInformation + vbOKOnly
        itxtChequeFinal.SetFocus
        Exit Sub
    End If
    
    If Val(itxtChequeIncial.Text) > Val(itxtChequeFinal.Text) Then
        MsgBox "Defina cheque Final debe ser mayor al cheque inicial", vbInformation + vbOKOnly
        itxtChequeFinal.SetFocus
        Exit Sub
    End If
    
    If Val(itxtChequeIncial.Text) <= Val(itxtChequeAntFinal.Text) Then
        MsgBox "El cheque inicial debe ser mayor al ultimo cheque registrado !!!", vbInformation + vbOKOnly
        itxtChequeIncial.SetFocus
        Exit Sub
    End If
    
    Dim oCuentaCheques As New CuentaCheques
    
    Call oCuentaCheques.chequesInsertaSerie(iCuentaCheques, _
                                            Val(itxtChequeIncial.Text), _
                                            Val(itxtChequeFinal.Text))
    
    cmdRegistrar.Enabled = False
    
    Call obtenRegistroActual(iCuentaCheques)
    itxtChequeIncial.Text = ""
    itxtChequeFinal.Text = ""
    
    Set oCuentaCheques = Nothing
    
    If lblCuentaCheques.Visible = True Then
        Unload Me
    End If
    
End Sub

Private Sub cmdActualizar_Click()

    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    Dim lRenglon As Long
    
    For lRenglon = 1 To sprCheques.DataRowCnt
        sprCheques.Row = lRenglon
        sprCheques.Col = 1
        
        If sprCheques.Text = "1" Then
            Set cCheque = New Collection
            'cCheque.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
            cCheque.Add oCampo.CreaCampo(adInteger, , , iCuentaCheques)
            sprCheques.Col = 2
            cCheque.Add oCampo.CreaCampo(adInteger, , , Val(sprCheques.Text))
            cCheque.Add oCampo.CreaCampo(adInteger, , , iEstadoNuevoCheque)
            cCheques.Add cCheque
        End If
        
    Next lRenglon
    
    Dim oCuentaCheques As New CuentaCheques
    Call oCuentaCheques.chequesaActualizaEstatusCobrado(cCheques)
    
    Call obtenEstatus(iCuentaCheques)
                                            
    Set oCuentaCheques = Nothing
    
End Sub

Private Sub cmdsalir_Click()
    iCuentaCheques = 0
    Unload Me
End Sub


