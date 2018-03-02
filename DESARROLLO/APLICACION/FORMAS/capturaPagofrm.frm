VERSION 5.00
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form capturaPagofrm 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Captura de Pagos"
   ClientHeight    =   9735
   ClientLeft      =   315
   ClientTop       =   30
   ClientWidth     =   9675
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdregistra 
      BackColor       =   &H80000003&
      Caption         =   "Registra Pagos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9150
      Width           =   1335
   End
   Begin Crystal.CrystalReport crPagos 
      Left            =   1590
      Top             =   9120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\facturas\pagos.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdgraba 
      BackColor       =   &H80000003&
      Caption         =   "Guardar Pagos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6810
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9150
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H80000003&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9150
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000003&
      Caption         =   "Pagos Capturados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   30
      TabIndex        =   0
      Top             =   3180
      Width           =   9615
      Begin FPSpread.vaSpread sprPago 
         Height          =   4425
         Left            =   30
         TabIndex        =   11
         Top             =   870
         Width           =   9525
         _Version        =   196608
         _ExtentX        =   16801
         _ExtentY        =   7805
         _StockProps     =   64
         BackColorStyle  =   2
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16306616
         GridColor       =   12092504
         MaxCols         =   8
         OperationMode   =   2
         SelectBlockOptions=   0
         ShadowColor     =   16306616
         SpreadDesigner  =   "capturaPagofrm.frx":0000
         Appearance      =   2
      End
      Begin VB.TextBox txttotadeudo 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   4
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin EditLib.fpCurrency txttotpagoold 
         Height          =   435
         Left            =   7350
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   5370
         Visible         =   0   'False
         Width           =   2205
         _Version        =   196608
         _ExtentX        =   3889
         _ExtentY        =   767
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483645
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
         ControlType     =   2
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   -1
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "99999"
         MinValue        =   "0"
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
      Begin VB.Label txttotpago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7380
         TabIndex        =   38
         Top             =   5355
         Width           =   2175
      End
      Begin VB.Label txtadeudo1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7290
         TabIndex        =   37
         Top             =   315
         Width           =   2220
      End
      Begin VB.Label txtpago1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4005
         TabIndex        =   36
         Top             =   315
         Width           =   2130
      End
      Begin VB.Label txtultimo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1665
         TabIndex        =   35
         Top             =   315
         Width           =   1500
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000003&
         Caption         =   "Pago:"
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
         Left            =   3240
         TabIndex        =   9
         Top             =   435
         Width           =   675
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000003&
         Caption         =   "Adeudo:"
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
         Left            =   6240
         TabIndex        =   8
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000003&
         Caption         =   "Último Folio:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   435
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000003&
         Caption         =   "Total Adeudos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2550
         TabIndex        =   3
         Top             =   5340
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000003&
         Caption         =   "Total Pagos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   2
         Top             =   5430
         Width           =   1515
      End
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   465
      Left            =   30
      TabIndex        =   10
      Top             =   2700
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   820
      _Version        =   196608
      ForeColor       =   192
      BackColor       =   -2147483645
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196608
      MinFontSize     =   1
      MaxFontSize     =   100
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2625
      Left            =   0
      TabIndex        =   13
      Top             =   60
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   4630
      _Version        =   196608
      BackColor       =   -2147483645
      Caption         =   "Crédito Datos"
      Begin MSMask.MaskEdBox txtpago 
         Height          =   375
         Left            =   3060
         TabIndex        =   28
         Top             =   630
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483645
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "$###,###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfolio 
         Height          =   375
         Left            =   135
         TabIndex        =   27
         Top             =   630
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483645
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdagregar 
         BackColor       =   &H80000003&
         Caption         =   "Agregar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8145
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   585
         Width           =   1395
      End
      Begin EditLib.fpLongInteger txtfolioold 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Visible         =   0   'False
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5027
         _ExtentY        =   661
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
         BackColor       =   -2147483645
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
         MaxValue        =   "49999"
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
      Begin EditLib.fpCurrency txtpagoold 
         Height          =   375
         Left            =   3030
         TabIndex        =   16
         Top             =   630
         Visible         =   0   'False
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
         _ExtentY        =   661
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
         BackColor       =   -2147483645
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
         MaxValue        =   "99999"
         MinValue        =   "0"
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
      Begin VB.Label txtcantatraso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5490
         TabIndex        =   33
         Top             =   1305
         Width           =   2400
      End
      Begin VB.Label txtcte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   135
         TabIndex        =   34
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label txtdias_atraso 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3060
         TabIndex        =   32
         Top             =   1305
         Width           =   2310
      End
      Begin VB.Label txtfecha 
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1410
         TabIndex        =   31
         Top             =   1305
         Width           =   1545
      End
      Begin VB.Label txtadeudo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5490
         TabIndex        =   30
         Top             =   630
         Width           =   2400
      End
      Begin VB.Label txtdias 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   29
         Top             =   1305
         Width           =   1185
      End
      Begin VB.Label txtnombre 
         BackColor       =   &H80000003&
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   1410
         TabIndex        =   26
         Top             =   2040
         Width           =   6465
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "Folio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   25
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3060
         TabIndex        =   24
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         Caption         =   "Adeudo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5490
         TabIndex        =   23
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000003&
         Caption         =   "Días de Crédito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   22
         Top             =   1095
         Width           =   1155
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000003&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   21
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000003&
         Caption         =   "Cantidad atrasada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5490
         TabIndex        =   20
         Top             =   1095
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000003&
         Caption         =   "Dias de atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3060
         TabIndex        =   19
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "Cliente número"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   18
         Top             =   1830
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Cliente nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1410
         TabIndex        =   17
         Top             =   1830
         Width           =   855
      End
   End
End
Attribute VB_Name = "capturaPagofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_FOLIO = 1
Private Const COL_NO_CLIENTE = 2
Private Const COL_NOMBRE_CLIENTE = 3
Private Const COL_FECHA = 4
Private Const COL_PAGO = 5
Private Const COL_ADEUDO = 6
Private Const COL_HORA = 7
Private Const COL_USUARIO = 8

Private lRowModifica As Long

Public nombreaux As String
Public strUsuario As String

Dim adeudo, pagogen As Double
'Dim bandera As Integer

Dim renglon As Long

Private Sub Form_Activate()
    
    txtdias.Caption = ""
    txtdias_atraso.Caption = ""
    txtcantatraso.Caption = ""

    txtfecha.Caption = Format(Now, "dd/mm/yyyy")
    
    'Si hay datos precapturados, muestralos
    If cPagosMem.Count > 0 Then
        
        fnLlenaTablaCollection sprPago, cPagosMem
        
        'Despliega el primer renglón
        sprPago.Row = 1
        sprPago.Col = COL_FOLIO
        txtultimo.Caption = sprPago.Text
        sprPago.Col = COL_PAGO
        txtpago1.Caption = sprPago.Text
        sprPago.Col = COL_ADEUDO
        txtadeudo1.Caption = sprPago.Text
        
        'Calcula el total cobrado
        txttotpago.Caption = Format(obtenTotalGrid(sprPago, COL_PAGO), "$#,####.00")
        
    End If

    If txtfolio.Enabled = True Then
        txtfolio.SetFocus
    End If
    
End Sub

'Private Sub Form_Load()
'
'    bandera = 1
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)

    If cmdgraba.Enabled = True Then
        
        If sprPago.DataRowCnt > 0 Then
            
            'Guarda los pagos precapturados
            obtenPagosMem
               
        End If
        
    End If
        
End Sub

Private Function obtenPagosMem()

    Dim cPagos As New Collection
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    Dim lRow As Long
    Dim lCol As Long
    Dim i As Integer
    
    If cPagosMem.Count > 0 Then
        For i = 1 To cPagosMem.Count
            cPagosMem.Remove 1
        Next
    End If
    
    For lRow = 1 To sprPago.DataRowCnt
    
        Set cRegistro = New Collection
        
        sprPago.Row = lRow
        
        sprPago.Col = COL_FOLIO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text)
        sprPago.Col = COL_NO_CLIENTE
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text)
        sprPago.Col = COL_NOMBRE_CLIENTE
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text)
        sprPago.Col = COL_FECHA
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text)
        sprPago.Col = COL_PAGO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPago.Text)))
        sprPago.Col = COL_ADEUDO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPago.Text)))
        sprPago.Col = COL_HORA
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text)
        
        'For lCol = COL_FOLIO To COL_HORA
        '
        '    sprPago.Col = lCol
        '    cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text)
        '
        'Next lCol
                
        cPagosMem.Add cRegistro
        
    Next lRow
        
End Function

Private Function despliegaDatosCredito(cCredito As Collection) As String

    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    Dim fCantidadTotal As Double
    Dim fCantidadPagada As Double
    Dim fCredito As Double
    Dim fFinanciamianto As Double
    Dim strNombreCliente As String
    Dim fechas, fechaini, pagado As Double
    Dim pagos As Integer
    Dim total, fAdeudo As Double
    Dim strMensaje As String
    
    despliegaDatosCredito = ""
    
    fechas = CDbl(CDate(Format(Now(), "dd/mm/yyyy")))
    
    Set cRegistro = cCredito(1)
    
    Set oCampo = cRegistro(10) 'Fecha
    If CDbl(Now) < CDbl(oCampo.Valor) Then
        
        despliegaDatosCredito = "La fecha de inicio de pago es mayor a hoy"
        Exit Function
        
    End If
    
    Set oCampo = cRegistro(3) 'Crédito
    If IsNull(oCampo.Valor) Then
        fCredito = 0#
    Else
        fCredito = oCampo.Valor
    End If
    
    Set oCampo = cRegistro(5) 'Financiamiento
    If IsNull(oCampo.Valor) Then
        fFinanciamianto = 0#
    Else
        fFinanciamianto = oCampo.Valor
    End If
    
    fCantidadTotal = fCredito + fFinanciamianto
    
    Set oCampo = cRegistro(21) 'Cantidad pagada
    fCantidadPagada = oCampo.Valor
    
    adeudo = fCantidadTotal - fCantidadPagada
    
    Set oCampo = cRegistro(2) 'Folio del crédito
    If fCantidadTotal - fCantidadPagada <= 0 Then
        despliegaDatosCredito = "¡La deuda correspondiente al documento " & oCampo.Valor & " ya esta liquidada!"
        'Exit Function
    End If
    
    Set oCampo = cRegistro(1) 'CLiente
    txtcte.Caption = oCampo.Valor
    
    Set oCampo = cRegistro(4) 'Cantidad a pagar
    txtpago.Text = Format(oCampo.Valor, "$#,####.00")
    pagogen = Format(oCampo.Valor, "$#,####.00")
    
    Set oCampo = cRegistro(6) 'Cantidad total
    total = oCampo.Valor
    
    Set oCampo = cRegistro(8) 'Fecha inicial del crédito
    fechaini = CDbl(oCampo.Valor)
    
    txtfecha.Caption = CDate(Format(Now, "dd/mm/yyyy"))
    
    Set oCampo = cRegistro(7) 'No de Pagos
    pagos = oCampo.Valor
    
    Set oCampo = cRegistro(10) 'Fecha contrato
    txtdias.Caption = CDate(Format(Now, "dd/mm/yyyy")) - CDate(oCampo.Valor) - 1
    
    Set oCampo = cRegistro(19) 'Nombre cliente
    strNombreCliente = oCampo.Valor
    Set oCampo = cRegistro(20) 'Apellido cliente
    txtNombre.Caption = strNombreCliente + " " + oCampo.Valor
    
    fAdeudo = fCantidadTotal - fCantidadPagada
    If fAdeudo > 0# Then
        cmdagregar.Enabled = True
    End If
    
    If fAdeudo > pagogen Then
        txtadeudo.Caption = Format(fAdeudo - pagogen, "$#,###.00")
    Else
        'txtadeudo.Text = Format(pagogen, "###,###,###,###0.00")
        txtadeudo.Caption = Format(0#, "$,####.00")
    End If
    
    pagado = fCantidadPagada
    
    txtdias_atraso.Caption = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado) / pagogen, "#,####.00")
    txtcantatraso.Caption = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado), "$#,####.00")

    txttotpago.Caption = Format(obtenTotalGrid(sprPago, COL_PAGO), "$#,####.00")
   
End Function

Private Function ObtenDatosCredito(lFactura As Long)

    If lFactura = 0 Then
        Exit Function
    End If
    
    Dim oPago As New Pago
    Dim lRowAnt As Long
    
    pnlMsg = ""
    If cmdagregar.Caption = "Agregar" And existeValorEnGrid(sprPago, COL_FOLIO, lFactura, lRowAnt) = True Then
        
        pnlMsg.Caption = "¡El No. de Folio " & lFactura & " ya se capturó, verifique por favor!" ', vbInformation + vbOKOnly
        
    Else
    
        If oPago.registrado(lFactura, Format(Now(), "dd/mm/yyyy")) Then
            
            pnlMsg.Caption = "¡Ya hay un pago registrado el día de hoy para el crédito con folio" & " " & lFactura & ", " & "verifique por favor!"
            txtfolio.SetFocus
            
        Else
        
            Dim oCredito As New credito
            
            If oCredito.datosCredito(lFactura) = False Then
            
                pnlMsg.Caption = "¡No hay un crédito con No. de folio " & lFactura & ", " & "verifique por favor!"
                txtfolio.SetFocus
                
            Else
                
                Dim strMsg As String
                strMsg = despliegaDatosCredito(oCredito.cDatos)
                
                If Len(strMsg) > 0 Then
                    pnlMsg = strMsg
                End If
                
                txtpago.SelStart = 0
                txtpago.SelLength = Len(txtpago.Text)
                txtpago.SetFocus
                                        
            End If
            
            Set oCredito = Nothing
            
        End If
        
        Set oPago = Nothing
    End If

End Function

Private Sub txtfolio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
                
        ObtenDatosCredito Val(txtfolio.Text)
        
    End If
    
End Sub

Private Sub txtfolio_LostFocus()
        ObtenDatosCredito Val(txtfolio.Text)
End Sub

Private Function lostFocusPago()
    
    If Val(txtfolio.Text) = 0 Then
        cmdagregar.Enabled = False
        txtfolio.SetFocus
    Else
    
        Dim fPago As Double
        fPago = Val(fnstrValor(txtpago.Text)) / 100#
        
        'If fPago >= 0# Then
            If adeudo > fPago Then
                txtadeudo.Caption = Format(adeudo - fPago, "$#,####.00")
            Else
                txtadeudo.Caption = Format(0, "$#,####.00")
            End If
            cmdagregar.Enabled = True
            cmdagregar.SetFocus
        'Else
        '    MsgBox "El monto del pago debe ser mayor a $0.0 peso", vbCritical, "Registro de Pagos"
        '    txtpago.SelLength = 255
        '    txtpago.SetFocus
        'End If
        'bandera = 1
    
        'cmdagregar_Click
    End If
    
End Function

Private Sub txtpago_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        lostFocusPago
    End If
    
End Sub

Private Sub txtpago_LostFocus()

    lostFocusPago
    
End Sub
'
'    Dim fPago As Double
'    fPago = Val(fnstrValor(txtpago.Text))
'
'    If fPago >= 0# Then
'        If adeudo > fPago Then
'            txtadeudo.Text = Format(adeudo - fPago, "###,###,###,###0.00")
'        Else
'            txtadeudo.Text = Format(adeudo - fPago, "###,###,###,###0.00")
'        End If
'    Else
'        MsgBox "El monto del pago debe ser mayor a $0.0 peso", vbCritical, "Registro de Pagos"
'        txtpago.SelLength = 255
'        txtpago.SetFocus
'    End If
'    bandera = 1
'
'End Sub

Private Function bajaCaptura(lRow As Long)

    Dim fPago As Double
    
    fPago = Val(txtpago.Text) / 100#
        
    'If fPago > 0# Then
        
        If cmdagregar.Caption = "Agregar" Then
        ' Insert one row before the specified row
            ' Increase the maximum number of rows by 1
            sprPago.MaxRows = sprPago.DataRowCnt + 1
            ' Specify a row
            sprPago.Row = lRow
            ' Insert a row
            sprPago.Action = SS_ACTION_INSERT_ROW
        Else
            sprPago.Row = lRow
        End If
    
        sprPago.Col = COL_FOLIO
        sprPago.Text = txtfolio.Text
        sprPago.Col = COL_NO_CLIENTE
        sprPago.Text = txtcte.Caption
        sprPago.Col = COL_NOMBRE_CLIENTE
        sprPago.Text = txtNombre.Caption
        sprPago.Col = COL_FECHA
        sprPago.Text = txtfecha.Caption
        sprPago.Col = COL_PAGO
        sprPago.Text = fPago 'Format(txtpago.Text, "###,###,###,###0.00")
        sprPago.Col = COL_ADEUDO
        sprPago.Text = txtadeudo.Caption
        sprPago.Col = COL_HORA
        sprPago.Text = Format(Now, "hh:mm AM/PM")
        sprPago.Col = COL_USUARIO
        sprPago.Text = nombreaux
        
        txttotpago.Caption = Format(obtenTotalGrid(sprPago, COL_PAGO), "$#,###.00")
        
        'bandera = 0
        
        txtdias.Caption = ""
        txtdias_atraso.Caption = ""
        txtcantatraso.Caption = ""
        txtcte.Caption = ""
        txtNombre.Caption = ""
        txtpago.Text = ""
        txtadeudo.Caption = ""
        
        txtfolio.Enabled = True
        txtfolio.Text = ""
        txtfolio.SetFocus

        cmdgraba.Enabled = True
        cmdagregar.Caption = "Agregar"
        cmdagregar.Enabled = False

    'Else
    '    MsgBox "¡El monto del pago debe ser mayor a $0.0 pesos, verifique por favor!", vbInformation + vbOKOnly, "Registro de Pagos"
    '    txtpago.Text = pagogen
    '    txtpago.SetFocus
    'End If
        
End Function

Private Sub cmdagregar_Click()

    If cmdagregar.Caption = "Modifica Pago" Then
        bajaCaptura renglon
    ElseIf cmdagregar.Caption <> "Modifica Pago" Then
        txtultimo.Caption = txtfolio.Text
        txtpago1.Caption = Format(Val(txtpago.Text) / 100#, "$#,####.00")
        txtadeudo1.Caption = txtadeudo.Caption
        bajaCaptura 1
    End If
    
End Sub

Private Sub sprPago_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row > 0 Then
        
        Dim totpago1, totadeudo1 As Double
        
        renglon = Row
        
        sprPago.Row = Row
        sprPago.Col = COL_FOLIO
        
        If sprPago.Text <> "" Then
        
            Dim oCredito As New credito
            Dim strMsg As String
            If oCredito.datosCredito(Val(sprPago.Text)) = False Then
            
                pnlMsg.Caption = "¡No hay un crédito con No. de folio " & sprPago.Text & ", " & "verifique por favor!"
                
            Else
                
                txtfolio.Text = sprPago.Text
                strMsg = despliegaDatosCredito(oCredito.cDatos)
                
                If Len(strMsg) > 0 Then
                    pnlMsg = strMsg
                End If
                                        
            End If
            
            Set oCredito = Nothing
    
            sprPago.Col = COL_PAGO
            totpago1 = CDbl(sprPago.Text)
            sprPago.Col = COL_ADEUDO
            totadeudo1 = CDbl(sprPago.Text)
            txttotpago.Caption = Format(CDbl(txttotpago.Caption) - totpago1, "###,###,###,###0.00")
                    
            cmdagregar.Caption = "Modifica Pago"
            txtfolio.Enabled = False
            txtpago.SetFocus
            
        End If
        
    End If
    
End Sub

Private Function obtenPagos(bRegistraGraba As Boolean) As Collection

    Dim cPagos As New Collection
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    Dim strUsuarioTemp As String
    Dim strHora As String
    Dim strLugar As String
    Dim iRegistraGraba As Integer
    Dim lRow As Long
    
    For lRow = 1 To sprPago.DataRowCnt
    
        Set cRegistro = New Collection
        
        sprPago.Row = lRow
        
        If bRegistraGraba = False Then
            strUsuarioTemp = "Nombre"
            strHora = ""
            strLugar = "captura"
            iRegistraGraba = 0
        Else
            strUsuarioTemp = nombreaux
            sprPago.Col = COL_HORA
            strHora = sprPago.Text
            strLugar = "cargaHH"
            iRegistraGraba = 1
        End If
        
        sprPago.Col = COL_FOLIO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(sprPago.Text))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , 0) 'No. de Pago
        cRegistro.Add oCampo.CreaCampo(adInteger, , , 69) 'Consecutivo
        sprPago.Col = COL_PAGO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPago.Text)))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , strUsuarioTemp) 'Usuario
        sprPago.Col = COL_FECHA
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text)
        sprPago.Col = COL_HORA
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text) 'Hora
        cRegistro.Add oCampo.CreaCampo(adInteger, , , strLugar) 'Lugar
        
        cRegistro.Add oCampo.CreaCampo(adInteger, , , lRow) 'Orden
        cRegistro.Add oCampo.CreaCampo(adInteger, , , iRegistraGraba) 'Graba = 1, Registra = 0
        
        sprPago.Col = COL_NO_CLIENTE
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(sprPago.Text))
        sprPago.Col = COL_ADEUDO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPago.Text)))
        
        cPagos.Add cRegistro
        
    Next lRow
    
    Set obtenPagos = cPagos
    
End Function

Private Sub fnImprime(strReporte As String, crObjeto As CrystalReport)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    Dim strNombreReporte As String
    
    cParametros.Add oCampo.CreaCampo(adInteger, , , 0)

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

Private Sub cmdgraba_Click()

    Dim oPago As New Pago
    
    If oPago.grabaPagos(obtenPagos(False)) Then
        
        MsgBox "Los pagos fueron registrados", vbInformation + vbOKOnly, "Registro de Pagos"
        cmdregistra.Enabled = True
        cmdgraba.Enabled = False
        fnImprime "pagos", crPagos
        sprPago.Enabled = False
        
        txtfolio.Enabled = False
        txtpago.Enabled = False
        
    End If

    Set oPago = Nothing

End Sub

Private Sub cmdregistra_Click()

    If MsgBox("Los pagos registrados están correctos?", vbQuestion + vbYesNo, "Registro de Pagos") = vbYes Then

        Dim oPago As New Pago
        
        If oPago.grabaPagos(obtenPagos(True)) Then
            
            MsgBox "Los pagos fueron registrados", vbInformation + vbOKOnly, "Registro de Pagos"
            cmdregistra.Enabled = False
            cmdgraba.Enabled = False
            
            fnLimpiaGrid sprPago
            txtfolio.Enabled = True
            txtpago.Enabled = True
            
        End If
        
        Set oPago = Nothing
        
    Else
        cmdgraba.Enabled = True
        cmdregistra.Enabled = False
    End If
    
    sprPago.Enabled = True
    
End Sub

Private Sub cmdsalir_Click()
    sicPrincipalfrm.Caption = ""
    sicPrincipalfrm.pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
    despliegaVentana portadafrm, WND_PORTADA
End Sub

