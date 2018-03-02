VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2C724BE0-A87B-11D1-8027-00A0C903B2B1}#6.0#0"; "TTFI6.ocx"
Begin VB.Form creditosfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel pnlCreditos 
      Height          =   5625
      Index           =   1
      Left            =   30
      TabIndex        =   2
      Top             =   450
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   9922
      _Version        =   196608
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salir"
         Height          =   435
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Del credito seleccionado, modifica pagos."
         Top             =   2730
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaldosIndividuales 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldos Individuales"
         Height          =   435
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Del credito seleccionado, modifica pagos."
         Top             =   2220
         Width           =   1335
      End
      Begin VB.CommandButton cmdReporteCreditosNuevos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reporte Créditos Nuevos"
         Height          =   435
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Del credito seleccionado, modifica pagos."
         Top             =   1770
         Width           =   1335
      End
      Begin Threed.SSCommand cmdPorCobrar 
         Height          =   465
         Left            =   11400
         TabIndex        =   23
         Top             =   4590
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   820
         _Version        =   196608
         BackColor       =   16777215
         PictureFrames   =   1
         Picture         =   "creditosfrm.frx":0000
         Caption         =   "Por Cobrar"
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin Crystal.CrystalReport crEspectativa 
         Left            =   10950
         Top             =   5040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Threed.SSPanel pnlF7 
         Height          =   165
         Left            =   7980
         TabIndex        =   17
         Top             =   90
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   291
         _Version        =   196608
         ForeColor       =   -2147483635
         BackColor       =   16777215
         Caption         =   "Modifica Crédito - Doble clik en el Folio deseado."
         BevelOuter      =   0
         Alignment       =   1
         RoundedCorners  =   0   'False
      End
      Begin FPSpread.vaSpread sprEspectativa 
         Height          =   2415
         Left            =   30
         TabIndex        =   16
         Top             =   3180
         Width           =   6585
         _Version        =   196608
         _ExtentX        =   11615
         _ExtentY        =   4260
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
         OperationMode   =   2
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         SpreadDesigner  =   "creditosfrm.frx":031A
      End
      Begin FPSpread.vaSpread sprCredito 
         Height          =   2325
         Left            =   30
         TabIndex        =   8
         Top             =   390
         Width           =   11325
         _Version        =   196608
         _ExtentX        =   19976
         _ExtentY        =   4101
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
         OperationMode   =   2
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   16777215
         SpreadDesigner  =   "creditosfrm.frx":1C69
      End
      Begin VB.CommandButton cmdModificaPagos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modifica Pagos"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Del credito seleccionado, modifica pagos."
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdcredito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nuevo crédito"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminaCredito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Elimina Crédito"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   810
         Width           =   1335
      End
      Begin VB.CommandButton cmdModificaCredito 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modifica Crédito"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   420
         Width           =   1335
      End
      Begin ComctlLib.TabStrip tabCredito 
         Height          =   2775
         Left            =   0
         TabIndex        =   7
         Top             =   -30
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   4895
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Vigentes"
               Key             =   ""
               Object.Tag             =   "V"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Pendientes"
               Key             =   ""
               Object.Tag             =   "P"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Cancelados"
               Key             =   ""
               Object.Tag             =   "C"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Vencidos"
               Key             =   ""
               Object.Tag             =   "E"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Terminados"
               Key             =   ""
               Object.Tag             =   "T"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   405
         Left            =   30
         TabIndex        =   18
         Top             =   2790
         Width           =   12885
         _ExtentX        =   22728
         _ExtentY        =   714
         _Version        =   196608
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Espectativa de cobranza y Deuda de creditos vigentes"
         RoundedCorners  =   0   'False
         Begin Crystal.CrystalReport crSaldos 
            Left            =   11160
            Top             =   -210
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
      End
      Begin EditLib.fpCurrency txttotpago 
         Height          =   405
         Left            =   9630
         TabIndex        =   19
         Top             =   4590
         Width           =   1725
         _Version        =   196608
         _ExtentX        =   3043
         _ExtentY        =   714
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   8388608
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
      Begin EditLib.fpCurrency txttotadeudo 
         Height          =   405
         Left            =   9630
         TabIndex        =   20
         Top             =   3750
         Width           =   1725
         _Version        =   196608
         _ExtentX        =   3043
         _ExtentY        =   714
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         ForeColor       =   8388608
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
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total Deuda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7500
         TabIndex        =   22
         Top             =   3750
         Width           =   1875
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total x Cobrar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7500
         TabIndex        =   21
         Top             =   4590
         Width           =   1815
      End
   End
   Begin Threed.SSPanel pnlCreditos 
      Height          =   5625
      Index           =   2
      Left            =   30
      TabIndex        =   1
      Top             =   450
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   9922
      _Version        =   196608
      RoundedCorners  =   0   'False
      Begin Threed.SSFrame SSFrame1 
         Height          =   825
         Left            =   1890
         TabIndex        =   11
         Top             =   270
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   1455
         _Version        =   196608
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Periodo"
         Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
            Height          =   345
            Left            =   5010
            TabIndex        =   12
            Top             =   390
            Width           =   1635
            _Version        =   65537
            _ExtentX        =   2884
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   -2147483633
         End
         Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo2 
            Height          =   345
            Left            =   1410
            TabIndex        =   13
            Top             =   390
            Width           =   1635
            _Version        =   65537
            _ExtentX        =   2884
            _ExtentY        =   609
            _StockProps     =   93
            BackColor       =   -2147483633
         End
         Begin VB.Label Label2 
            Caption         =   "Fin"
            Height          =   315
            Left            =   3900
            TabIndex        =   15
            Top             =   420
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Inicio"
            Height          =   315
            Left            =   150
            TabIndex        =   14
            Top             =   420
            Width           =   1245
         End
      End
      Begin VtChartLib6.VtChart VtChart1 
         Height          =   3645
         Left            =   3600
         TabIndex        =   10
         Top             =   1410
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   6429
         _0              =   $"creditosfrm.frx":37A9
         _1              =   $"creditosfrm.frx":3BB2
         _2              =   $"creditosfrm.frx":3FBB
         _3              =   $"creditosfrm.frx":43C4
         _4              =   $"creditosfrm.frx":47CD
         _5              =   $"creditosfrm.frx":4BD6
         _6              =   $"creditosfrm.frx":4FDF
         _7              =   $"creditosfrm.frx":53E8
         _8              =   $"creditosfrm.frx":57F1
         _9              =   $"creditosfrm.frx":5BFA
         _10             =   $"creditosfrm.frx":6003
         _11             =   $"creditosfrm.frx":640C
         _12             =   $"creditosfrm.frx":6815
         _13             =   $"creditosfrm.frx":6C1E
         _14             =   "-@.??????.D@@@@@D@@@D@@@D@@@)]A@@c@@@G@)V4)L)i34@B@C@@@@@D@@@@@A@@@@@.??????????????.A@@@@@d@@@I@)V4-Window-@C@D@@@@@,E9B0"
         _count          =   15
         _ver            =   2
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3585
         Left            =   330
         TabIndex        =   9
         Top             =   1410
         Width           =   2235
         _Version        =   196608
         _ExtentX        =   3942
         _ExtentY        =   6324
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
         MaxCols         =   2
         SpreadDesigner  =   "creditosfrm.frx":7027
      End
   End
   Begin ComctlLib.TabStrip tabOperacionProyeccion 
      Height          =   6075
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12915
      _ExtentX        =   22781
      _ExtentY        =   10716
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Operación"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Proyección"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "creditosfrm.frx":8836
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196608
      MinFontSize     =   1
      MaxFontSize     =   100
   End
End
Attribute VB_Name = "creditosfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CREDITO_FACTURA = 1

Private Const COL_FOLIO = 1

Private Const COL_ESPECTATIVA_FOLIO = 1
Private Const COL_ESPECTATIVA_ULITMO_PAGO = 2
Private Const COL_ESPECTATIVA_ADEUDO = 3
Private Const COL_ESPECTATIVA_PAGO = 4
Private Const COL_ESPECTATIVA_DIAS_CREDITO = 5


Private iTabActivo As Integer

Private miFrameActivoOP As Integer

Private Sub cmdReporteCreditosNuevos_Click()
    reCreditosfrm.Show vbModal
End Sub

Private Sub cmdSaldosIndividuales_Click()

    If sprCredito.DataRowCnt > 0 Then
        
        
        sprCredito.Col = 1
        sprCredito.Row = sprCredito.ActiveRow
        
        Dim oReporte As New Reporte
        Dim cParametros As New Collection
        Dim oCampo As New Campo
        
        cParametros.Add oCampo.CreaCampo(adInteger, , , Val(sprCredito.Text))
        
        oReporte.oCrystalReport = crSaldos
        If gstrReporteEnPantalla = "Si" Then
            oReporte.bVistaPreliminar = True
        Else
            oReporte.bVistaPreliminar = False
        End If
        oReporte.strImpresora = gPrintPed
        oReporte.strNombreReporte = DirSys & "saldosIndividuales.rpt"
        oReporte.cParametros = cParametros
        oReporte.fnImprime
        Set oReporte = Nothing
        
'        saldosfrm.iFolio = sprCredito.Text
'        saldosfrm.Show vbModal
    
    End If
    
End Sub

Private Sub cmdsalir_Click()
    
    sicPrincipalfrm.Caption = ""
    sicPrincipalfrm.pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
    despliegaVentana portadafrm, WND_PORTADA
    
End Sub

Private Sub Form_Activate()
    miFrameActivoOP = 1
End Sub

Private Sub Form_Load()
    
    Dim oCredito As New credito
    
    If oCredito.obtenPorEstatus("V", Format(Now(), "dd/mm/yyyy")) = True Then
        Call fnLlenaTablaCollection(sprCredito, oCredito.cDatos)
    End If
    
    If oCredito.obtenEspectativa(Format(Now(), "dd/mm/yyyy")) = True Then
        Call fnLlenaTablaCollection(sprEspectativa, oCredito.cDatos)
        txttotadeudo.Text = obtenTotalGrid(sprEspectativa, COL_ESPECTATIVA_ADEUDO)
        txttotpago.Text = obtenTotalGrid(sprEspectativa, COL_ESPECTATIVA_PAGO)
    End If
    
    Set oCredito = Nothing
    
    cmdModificaCredito.Enabled = True
    cmdModificaPagos.Enabled = False
    cmdEliminaCredito.Enabled = True
       
End Sub

Private Sub tabCredito_Click()

    If tabCredito.SelectedItem.Index = iTabActivo Then Exit Sub ' No need to change frame.

    Dim oCredito As New credito
    
    If oCredito.obtenPorEstatus(tabCredito.SelectedItem.Tag, Format(Now(), "dd/mm/yyyy")) = True Then
        Call fnLlenaTablaCollection(sprCredito, oCredito.cDatos)
    End If
    
    Select Case tabCredito.SelectedItem.Tag
        Case Is = "E"
            cmdModificaCredito.Enabled = True
            cmdModificaPagos.Enabled = False
            cmdEliminaCredito.Enabled = True
        Case Is = "V"
            cmdModificaCredito.Enabled = True
            cmdModificaPagos.Enabled = True
            cmdEliminaCredito.Enabled = False
        Case Is = "P"
            cmdModificaCredito.Enabled = True
            cmdModificaPagos.Enabled = False
            cmdEliminaCredito.Enabled = False
        Case Is = "C"
            cmdModificaCredito.Enabled = False
            cmdModificaPagos.Enabled = False
            cmdEliminaCredito.Enabled = True
        Case Is = "T"
            cmdModificaCredito.Enabled = False
            cmdModificaPagos.Enabled = False
            cmdEliminaCredito.Enabled = True
            
    End Select
    
    iTabActivo = tabCredito.SelectedItem.Index
    Set oCredito = Nothing

End Sub

Private Sub sprCredito_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Call cmdModificaCredito_Click

End Sub

Private Sub cmdModificaCredito_Click()

    Select Case tabCredito.SelectedItem.Tag
        Case Is = "E", "V", "P"
            
            sprCredito.Row = sprCredito.ActiveRow
            sprCredito.Col = CREDITO_FACTURA
            If Len(sprCredito.Text) > 0 Then
                modificaCreditofrm.iNoFactura = Val(sprCredito.Text)
                modificaCreditofrm.Show vbModal
            End If
        
    End Select

End Sub

Private Sub cmdModificaPagos_Click()

    accesofrm.Show vbModal

    If accesofrm.bPermiteAcceso = True Then
    
        sprCredito.Col = COL_FOLIO
        sprCredito.Row = sprCredito.ActiveRow
        modificaPagofrm.lFolio = Val(sprCredito.Text)
        
        modificaPagofrm.Show vbModal
    
    End If
    
End Sub

Private Sub cmdEliminaCredito_Click()
    
    Select Case tabCredito.SelectedItem.Tag
        
        Case Is = "T", "C", "E"
            
            accesofrm.Show vbModal
        
            If accesofrm.bPermiteAcceso = True Then
                
                sprCredito.Row = sprCredito.ActiveRow
                sprCredito.Col = CREDITO_FACTURA
                If Len(sprCredito.Text) > 0 Then
                    
                    Dim oCredito As New credito
                    
                    oCredito.elimina (Val(sprCredito.Text))
                                                            
                    If oCredito.obtenPorEstatus(tabCredito.SelectedItem.Tag, Format(Now(), "dd/mm/yyyy")) = True Then
                        Call fnLlenaTablaCollection(sprCredito, oCredito.cDatos)
                    End If
                    
                    Set oCredito = Nothing
                    
                End If
                
            End If
        
    End Select

End Sub

Private Sub cmdPorCobrar_Click()
    fnImprime (Format(Now(), "dd/mm/yyyy"))
End Sub

Private Sub tabOperacionProyeccion_Click()
    If tabOperacionProyeccion.SelectedItem.Index = miFrameActivoOP Then Exit Sub ' No need to change frame.
    
    ' Comosea, oculta el frame anterior, muestra el nuevo.
    pnlCreditos(tabOperacionProyeccion.SelectedItem.Index).Visible = True
    pnlCreditos(miFrameActivoOP).Visible = False
    
    miFrameActivoOP = tabOperacionProyeccion.SelectedItem.Index

End Sub

Private Sub fnImprime(strFecha As String)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    Dim strNombreReporte As String
    
    strNombreReporte = "espectativa.rpt"
    
    oReporte.oCrystalReport = crEspectativa
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


