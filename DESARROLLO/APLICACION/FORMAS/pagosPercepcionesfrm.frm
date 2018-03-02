VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form pagosPercepcionesfrm 
   BorderStyle     =   0  'None
   Caption         =   "Pagos"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   90
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   12960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlPagos 
      Height          =   4425
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   510
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   7805
      _Version        =   196609
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPSpread.vaSpread sprPagosVencidos 
         Height          =   2325
         Left            =   0
         TabIndex        =   0
         Top             =   390
         Width           =   10305
         _Version        =   196608
         _ExtentX        =   18177
         _ExtentY        =   4101
         _StockProps     =   64
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
         MaxCols         =   11
         SpreadDesigner  =   "pagosPercepcionesfrm.frx":0000
         VisibleCols     =   5
      End
   End
   Begin Threed.SSPanel pnlPagos 
      Height          =   4425
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   510
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   7805
      _Version        =   196609
      Caption         =   "SSPanel1"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cbPeriodo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3180
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   30
         Width           =   4065
      End
      Begin FPSpread.vaSpread sprPagos 
         Height          =   2325
         Left            =   0
         TabIndex        =   2
         Top             =   390
         Width           =   10305
         _Version        =   196608
         _ExtentX        =   18177
         _ExtentY        =   4101
         _StockProps     =   64
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
         SpreadDesigner  =   "pagosPercepcionesfrm.frx":1B27
         VisibleCols     =   5
      End
      Begin VB.Label Label2 
         Caption         =   "Muestra los pagos de los próximos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "días."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7230
         TabIndex        =   7
         Top             =   30
         Width           =   495
      End
   End
   Begin Threed.SSPanel pnlPagos 
      Height          =   4425
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   510
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   7805
      _Version        =   196609
      Caption         =   "SSPanel2"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox cbProveedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1500
         TabIndex        =   21
         Top             =   510
         Width           =   3225
      End
      Begin FPSpread.vaSpread sprPagosRealizados 
         Height          =   2325
         Left            =   30
         TabIndex        =   20
         Top             =   1140
         Width           =   10305
         _Version        =   196608
         _ExtentX        =   18177
         _ExtentY        =   4101
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         SpreadDesigner  =   "pagosPercepcionesfrm.frx":356A
         VisibleCols     =   5
      End
      Begin SSCalendarWidgets_A.SSDateCombo fFlujoCaja 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   3
         EndProperty
         Height          =   405
         Left            =   5880
         TabIndex        =   23
         Top             =   480
         Width           =   1755
         _Version        =   65537
         _ExtentX        =   3096
         _ExtentY        =   714
         _StockProps     =   93
         ForeColor       =   -2147483630
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
      Begin SSCalendarWidgets_A.SSDateCombo fFlujoCajaFin 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   3
         EndProperty
         Height          =   435
         Left            =   8550
         TabIndex        =   24
         Top             =   480
         Width           =   1725
         _Version        =   65537
         _ExtentX        =   3043
         _ExtentY        =   767
         _StockProps     =   93
         ForeColor       =   -2147483630
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Del día:"
         Height          =   465
         Left            =   4830
         TabIndex        =   26
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Al día:"
         Height          =   465
         Left            =   7620
         TabIndex        =   25
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Proveedor:"
         Height          =   405
         Left            =   60
         TabIndex        =   22
         Top             =   510
         Width           =   1665
      End
   End
   Begin EditLib.fpCurrency txtfTotalVencido 
      Height          =   495
      Left            =   10980
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1755
      _Version        =   196608
      _ExtentX        =   3096
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483634
      ForeColor       =   255
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
      AlignTextV      =   0
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
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ","
      UseSeparator    =   -1  'True
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
   Begin VB.CommandButton cmdRegistrarPago 
      Caption         =   "Realizar pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7620
      TabIndex        =   4
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
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
      Height          =   465
      Left            =   9120
      TabIndex        =   6
      Top             =   5025
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Programar Nuevo Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6360
      TabIndex        =   3
      Top             =   5025
      Width           =   1215
   End
   Begin EditLib.fpDateTime dtFecha 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4950
      Width           =   1515
      _Version        =   196608
      _ExtentX        =   2672
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      AlignTextH      =   0
      AlignTextV      =   0
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
      Text            =   "04/09/2007"
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
      ThreeDFrameColor=   -2147483633
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
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtfTotalProximosPagos 
      Height          =   495
      Left            =   10980
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2310
      Width           =   1755
      _Version        =   196608
      _ExtentX        =   3096
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483634
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
      AlignTextV      =   0
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
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ","
      UseSeparator    =   -1  'True
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
   Begin EditLib.fpCurrency txtfSaldoBancos 
      Height          =   495
      Left            =   10980
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1755
      _Version        =   196608
      _ExtentX        =   3096
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483634
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
      AlignTextV      =   0
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
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ","
      UseSeparator    =   -1  'True
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
   Begin EditLib.fpCurrency txtfBalanceTotal 
      Height          =   495
      Left            =   10980
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1755
      _Version        =   196608
      _ExtentX        =   3096
      _ExtentY        =   873
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483634
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
      AlignTextV      =   0
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
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ","
      UseSeparator    =   -1  'True
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
   Begin ComctlLib.TabStrip tabPagos 
      Height          =   4785
      Left            =   0
      TabIndex        =   27
      Top             =   150
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   8440
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Vencidos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Próximos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Realizados"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Total:"
      Height          =   345
      Left            =   10410
      TabIndex        =   18
      Top             =   4020
      Width           =   1875
   End
   Begin VB.Label Label5 
      Caption         =   "Bancos:"
      Height          =   345
      Left            =   10410
      TabIndex        =   17
      Top             =   3000
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Próximos:"
      Height          =   345
      Left            =   10410
      TabIndex        =   16
      Top             =   1950
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Vencido:"
      Height          =   345
      Left            =   10410
      TabIndex        =   15
      Top             =   930
      Width           =   1875
   End
End
Attribute VB_Name = "pagosPercepcionesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const COL_CONCEPTO = 1
Const COL_MONTO = 2
Const COL_FECHA_PAGO = 3
Const COL_FRECUENCIA = 4
Const COL_FORMA_PAGO = 5
Const COL_CTA_CONTABLE = 6
Const COL_CTA_CHEQUES = 7
Const COL_PAGO = 8
Const COL_PAGO_CONSECUTIVO = 9
Const COL_FOLIO_DOCUMENTO = 10
Const COL_PAGO_CONCEPTO = 11

Const COL_PAGADO_CONCEPTO = 1
Const COL_PAGADO_MONTO = 2
Const COL_PAGADO_FECHA_PAGO = 3
Const COL_PAGADO_FORMA_PAGO = 4
Const COL_PAGADO_CTA_CONTABLE = 5
Const COL_PAGADO_CTA_CHEQUES = 6
Const COL_PAGADO_PAGO = 7
Const COL_PAGADO_PAGO_CONSECUTIVO = 89
Const COL_PAGADO_FOLIO_DOCUMENTO = 9
Const COL_PAGADO_PAGO_CONCEPTO = 10

Const PAGO_VENCIDO = 0
Const PAGO_NORMAL = 1
Const CALENDARIO_PAGO = 2

Private miFrameActivo As Integer

Private Sub Form_Load()
    Dim oPago As New cProveedor

    If oPago.periodicidadDias() Then
        Call fnLlenaComboCollecion(cbPeriodo, oPago.cDatos, 0, "")
    End If
    Set oPago = Nothing

    Call despligaPagosSaldos
    
End Sub

Private Function despligaPagosSaldos()

    'Carga catálogo de periodos de pago
    Dim dTotalVencido As Double
    Dim dTotalProximosPagos As Double
    Dim dSaldoBancos As Double
    Dim dtFechaFinal As String
    
    Dim oPago As New cProveedor
        
    dtFecha.Text = Date
    If cbPeriodo.ListIndex = -1 Then
        dtFechaFinal = dtFecha.Text
    Else
        dtFechaFinal = DateAdd("d", Val(cbPeriodo.Text), dtFecha.Text)
    End If
    
    Call fnLimpiaGrid(sprPagos)
    If oPago.obtenPagos(dtFecha.Text, dtFechaFinal) = True Then
        Call fnLlenaTablaCollection(sprPagos, oPago.cDatos)
    End If
    
    Call fnLimpiaGrid(sprPagosVencidos)
    If oPago.obtenPagosVencidos(dtFecha.Text) = True Then
        Call fnLlenaTablaCollection(sprPagosVencidos, oPago.cDatos)
    End If
    
    Set oPago = Nothing
    
    'CALCULA TOTALES
    txtfTotalVencido.Text = Format(obtenTotalGrid(sprPagosVencidos, COL_MONTO), "$#,###.00")
    txtfTotalProximosPagos.Text = Format(obtenTotalGrid(sprPagos, COL_MONTO), "$#,###.00")
    txtfSaldoBancos.Text = Format(obtenSaldoBancos(), "$#,###.00")
    dTotalVencido = Val(fnstrValor(txtfTotalVencido.Text))
    dTotalProximosPagos = Val(fnstrValor(txtfTotalProximosPagos.Text))
    dSaldoBancos = Val(fnstrValor(txtfSaldoBancos.Text))
    txtfBalanceTotal.Text = Format(Abs(dSaldoBancos - (dTotalVencido + dTotalProximosPagos)), "$#,###.00")
    
    If dSaldoBancos - (dTotalVencido + dTotalProximosPagos) > 0 Then
        txtfBalanceTotal.ForeColor = &H8000&     'RGB(0, 255, 255)
    Else
        txtfBalanceTotal.ForeColor = RGB(255, 0, 0)
    End If

End Function

'Private Function obtenSaldoBancos(iSalon As Integer) As Double
Private Function obtenSaldoBancos() As Double
    
    Dim oCuentaCheques As New CuentaCheques
    
    obtenSaldoBancos = oCuentaCheques.saldoBancos()
    Set oCuentaCheques = Nothing
    
End Function

Private Sub fFlujoCaja_Click()
    Consulta
End Sub

Private Sub fFlujoCaja_Change()
    Consulta
End Sub

Private Sub fFlujoCaja_CloseUp()
    Consulta
End Sub

Private Sub fFlujoCaja_Spin(OldDate As String, NewDate As String)
    Consulta
End Sub

Private Sub fFlujoCajaFin_Click()
    Consulta
End Sub

Private Sub fFlujoCajaFin_Change()
    Consulta
End Sub
Private Sub fFlujoCajaFin_CloseUp()
    Consulta
End Sub

Private Sub fFlujoCajaFin_Spin(OldDate As String, NewDate As String)
    Consulta
End Sub

Private Function Consulta()

    If 0 <= DateDiff("d", CDate(fFlujoCaja.Text), CDate(fFlujoCajaFin.Text)) Then
        Dim oPago As New cProveedor
        Call fnLimpiaGrid(sprPagosRealizados)
        If oPago.pagosRealizados(gAlmacen, fFlujoCaja.Text, fFlujoCajaFin.Text) = True Then
            Call fnLlenaTablaCollection(sprPagosRealizados, oPago.cDatos)
        End If
        
        Set oPago = Nothing
        
    Else
        MsgBox "La fecha final es mayor a la inicial, verifique periodo. !!! " 'cmdConsulta.Enabled = False
    End If
    
End Function

Private Sub cbPeriodo_Click()
    
    If cbPeriodo.ListIndex = -1 Then
        MsgBox "Seleccione por favor los dias que desea ver!! ", vbInformation + vbOKOnly
        Exit Sub
    Else
        Dim dTotalVencido As Double
        Dim dTotalProximosPagos As Double
        Dim dSaldoBancos As Double
    
        Dim dtFechaFinal As String
        Dim oPago As New cProveedor
        
        dtFechaFinal = DateAdd("d", Val(cbPeriodo.Text), dtFecha.Text)
        
        Call fnLimpiaGrid(sprPagos)
        If oPago.obtenPagos(dtFecha.Text, dtFechaFinal) = True Then
            Call fnLlenaTablaCollection(sprPagos, oPago.cDatos)
        End If
        
        Set oPago = Nothing
        
        txtfTotalVencido = obtenTotalGrid(sprPagosVencidos, COL_MONTO)
        txtfTotalProximosPagos = obtenTotalGrid(sprPagos, COL_MONTO)
        txtfSaldoBancos = obtenSaldoBancos()
        dTotalVencido = Val(fnstrValor(txtfTotalVencido.Text))
        dTotalProximosPagos = Val(fnstrValor(txtfTotalProximosPagos.Text))
        dSaldoBancos = Val(fnstrValor(txtfSaldoBancos.Text))
        txtfBalanceTotal.Text = Format(Abs(dSaldoBancos - (dTotalVencido + dTotalProximosPagos)), "$#,###.00")
        
        If dSaldoBancos - (dTotalVencido + dTotalProximosPagos) > 0 Then
            txtfBalanceTotal.ForeColor = &H8000& 'RGB(0, 255, 100)
        Else
            txtfBalanceTotal.ForeColor = RGB(255, 0, 0)
        End If
        
    End If

End Sub

Private Sub cmdNuevo_Click()
    pagofrm.bCambio = False
    pagofrm.Show vbModal
    'Actualiza el desplegado
    Call despligaPagosSaldos

End Sub

Private Sub cmdRegistrarPago_Click()
    
    Select Case miFrameActivo
        Case Is = PAGO_NORMAL
        
            'Toma los datos del reglon actual y pasalos a la forma
            sprPagos.Row = sprPagos.ActiveRow
            sprPagos.Col = COL_CONCEPTO
            
            If sprPagos.Text <> "" Then
                
                registroPagofrm.strConcepto = sprPagos.Text
                
                sprPagos.Col = COL_CTA_CHEQUES
                registroPagofrm.iCuentaBanco = sprPagos.Text
                
                sprPagos.Col = COL_MONTO
                registroPagofrm.dMonto = sprPagos.Text
                
                sprPagos.Col = COL_FECHA_PAGO
                registroPagofrm.strFecha = sprPagos.Text
                
                sprPagos.Col = COL_CTA_CONTABLE
                If sprPagos.Text = "" Then
                    registroPagofrm.iCuentaContable = 0
                Else
                    registroPagofrm.iCuentaContable = sprPagos.Text
                End If
                
                sprPagos.Col = COL_PAGO
                registroPagofrm.iPago = sprPagos.Text
                
                sprPagos.Col = COL_PAGO_CONSECUTIVO
                registroPagofrm.iPagoConsecutivo = sprPagos.Text
                
                sprPagos.Col = COL_PAGO_CONCEPTO
                registroPagofrm.iConcepto = sprPagos.Text
                
                registroPagofrm.Show vbModal
                
                'Registrado el pago, actualizar la vista
                Call despligaPagosSaldos
                
            Else
                MsgBox "Seleccione el pago que desea realizar !!!", vbInformation + vbOKOnly
            End If
        Case Is = PAGO_VENCIDO
            
            'Toma los datos del reglon actual y pasalos a la forma
            sprPagosVencidos.Row = sprPagosVencidos.ActiveRow
            sprPagosVencidos.Col = COL_CONCEPTO
            
            If sprPagosVencidos.Text <> "" Then
                
                registroPagofrm.strConcepto = sprPagosVencidos.Text
                
                sprPagosVencidos.Col = COL_CTA_CHEQUES
                registroPagofrm.iCuentaBanco = sprPagosVencidos.Text
                
                sprPagosVencidos.Col = COL_MONTO
                registroPagofrm.dMonto = sprPagosVencidos.Text
                
                sprPagosVencidos.Col = COL_FECHA_PAGO
                registroPagofrm.strFecha = sprPagosVencidos.Text
                
                sprPagosVencidos.Col = COL_CTA_CONTABLE
                
                registroPagofrm.iCuentaContable = 0
                
                sprPagosVencidos.Col = COL_PAGO
                registroPagofrm.iPago = sprPagosVencidos.Text
                
                sprPagosVencidos.Col = COL_PAGO_CONSECUTIVO
                registroPagofrm.iPagoConsecutivo = sprPagosVencidos.Text
                
                sprPagosVencidos.Col = COL_PAGO_CONCEPTO
                registroPagofrm.iConcepto = sprPagosVencidos.Text
                
                registroPagofrm.Show vbModal
    
                Call despligaPagosSaldos
                
            Else
                MsgBox "Seleccione el pago que desea realizar !!!", vbInformation + vbOKOnly
            End If
    End Select
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub sprPagos_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call cmdRegistrarPago_Click
End Sub

Private Sub sprPagosVencidos_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call cmdRegistrarPago_Click
End Sub

Private Sub tabPagos_Click()
    If tabPagos.SelectedItem.Index - 1 = miFrameActivo Then Exit Sub ' No need to change frame.
    
    ' Comosea, oculta el frame anterior, muestra el nuevo.
    pnlPagos(tabPagos.SelectedItem.Index - 1).Visible = True
    pnlPagos(miFrameActivo).Visible = False
    
    miFrameActivo = tabPagos.SelectedItem.Index - 1

End Sub
