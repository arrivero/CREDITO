VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2C724BE0-A87B-11D1-8027-00A0C903B2B1}#6.0#0"; "TTFI6.ocx"
Begin VB.Form resumenGraficofrm 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel pnlGraficos 
      Height          =   5385
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   9499
      _Version        =   196608
      BackColor       =   -2147483634
      Caption         =   "1"
      RoundedCorners  =   0   'False
      Begin Threed.SSOption rbDineroCtesCobrados 
         Height          =   225
         Index           =   0
         Left            =   4710
         TabIndex        =   3
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   397
         _Version        =   196608
         BackColor       =   -2147483634
         Caption         =   "Dinero Cobrado"
         Value           =   -1
      End
      Begin VB.ComboBox cbPeriodo 
         Height          =   315
         ItemData        =   "resumenGraficofrm.frx":0000
         Left            =   4740
         List            =   "resumenGraficofrm.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   2565
      End
      Begin VtChartLib6.VtChart grResumen 
         Height          =   4035
         Left            =   60
         TabIndex        =   9
         Top             =   1020
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   7117
         _0              =   $"resumenGraficofrm.frx":0052
         _1              =   $"resumenGraficofrm.frx":045C
         _2              =   $"resumenGraficofrm.frx":0865
         _3              =   $"resumenGraficofrm.frx":0C6E
         _4              =   $"resumenGraficofrm.frx":1077
         _5              =   $"resumenGraficofrm.frx":1480
         _6              =   $"resumenGraficofrm.frx":1889
         _7              =   $"resumenGraficofrm.frx":1C92
         _8              =   $"resumenGraficofrm.frx":209B
         _9              =   $"resumenGraficofrm.frx":24A5
         _10             =   $"resumenGraficofrm.frx":28AE
         _11             =   $"resumenGraficofrm.frx":2CB7
         _12             =   $"resumenGraficofrm.frx":30C0
         _13             =   $"resumenGraficofrm.frx":34C9
         _14             =   $"resumenGraficofrm.frx":38D2
         _15             =   $"resumenGraficofrm.frx":3CDB
         _16             =   $"resumenGraficofrm.frx":40E4
         _17             =   $"resumenGraficofrm.frx":44ED
         _18             =   $"resumenGraficofrm.frx":48F6
         _19             =   $"resumenGraficofrm.frx":4CFF
         _20             =   $"resumenGraficofrm.frx":5108
         _21             =   $"resumenGraficofrm.frx":5511
         _22             =   $"resumenGraficofrm.frx":591A
         _23             =   $"resumenGraficofrm.frx":5D23
         _24             =   $"resumenGraficofrm.frx":612C
         _25             =   $"resumenGraficofrm.frx":6535
         _26             =   $"resumenGraficofrm.frx":693E
         _27             =   $"resumenGraficofrm.frx":6D47
         _28             =   $"resumenGraficofrm.frx":7150
         _29             =   $"resumenGraficofrm.frx":755A
         _30             =   $"resumenGraficofrm.frx":7963
         _31             =   $"resumenGraficofrm.frx":7D6C
         _32             =   $"resumenGraficofrm.frx":8175
         _33             =   $"resumenGraficofrm.frx":857E
         _34             =   $"resumenGraficofrm.frx":8987
         _35             =   $"resumenGraficofrm.frx":8D90
         _36             =   $"resumenGraficofrm.frx":9199
         _37             =   $"resumenGraficofrm.frx":95A2
         _38             =   $"resumenGraficofrm.frx":99AB
         _39             =   $"resumenGraficofrm.frx":9DB4
         _40             =   $"resumenGraficofrm.frx":A1BD
         _41             =   $"resumenGraficofrm.frx":A5C6
         _42             =   $"resumenGraficofrm.frx":A9CF
         _43             =   $"resumenGraficofrm.frx":ADD8
         _44             =   $"resumenGraficofrm.frx":B1E1
         _45             =   $"resumenGraficofrm.frx":B5EA
         _46             =   $"resumenGraficofrm.frx":B9F4
         _47             =   $"resumenGraficofrm.frx":BDFE
         _48             =   $"resumenGraficofrm.frx":C207
         _49             =   $"resumenGraficofrm.frx":C610
         _50             =   $"resumenGraficofrm.frx":CA19
         _51             =   $"resumenGraficofrm.frx":CE22
         _52             =   $"resumenGraficofrm.frx":D22B
         _53             =   $"resumenGraficofrm.frx":D634
         _54             =   $"resumenGraficofrm.frx":DA3D
         _55             =   $"resumenGraficofrm.frx":DE46
         _56             =   $"resumenGraficofrm.frx":E24F
         _57             =   $"resumenGraficofrm.frx":E659
         _58             =   $"resumenGraficofrm.frx":EA62
         _59             =   $"resumenGraficofrm.frx":EE6B
         _60             =   $"resumenGraficofrm.frx":F274
         _61             =   $"resumenGraficofrm.frx":F67D
         _62             =   $"resumenGraficofrm.frx":FA86
         _63             =   $"resumenGraficofrm.frx":FE8F
         _64             =   $"resumenGraficofrm.frx":10298
         _65             =   $"resumenGraficofrm.frx":106A1
         _66             =   $"resumenGraficofrm.frx":10AAA
         _67             =   $"resumenGraficofrm.frx":10EB3
         _68             =   $"resumenGraficofrm.frx":112BD
         _69             =   $"resumenGraficofrm.frx":116C6
         _70             =   $"resumenGraficofrm.frx":11ACF
         _71             =   $"resumenGraficofrm.frx":11ED8
         _72             =   $"resumenGraficofrm.frx":122E1
         _73             =   $"resumenGraficofrm.frx":126EA
         _74             =   $"resumenGraficofrm.frx":12AF3
         _75             =   $"resumenGraficofrm.frx":12EFC
         _76             =   $"resumenGraficofrm.frx":13305
         _77             =   $"resumenGraficofrm.frx":1370E
         _78             =   $"resumenGraficofrm.frx":13B17
         _79             =   $"resumenGraficofrm.frx":13F20
         _80             =   $"resumenGraficofrm.frx":14329
         _81             =   $"resumenGraficofrm.frx":14732
         _82             =   $"resumenGraficofrm.frx":14B3B
         _83             =   $"resumenGraficofrm.frx":14F44
         _84             =   $"resumenGraficofrm.frx":1534D
         _85             =   $"resumenGraficofrm.frx":15756
         _86             =   $"resumenGraficofrm.frx":15B5F
         _count          =   87
         _ver            =   2
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   975
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1720
         _Version        =   196608
         BackColor       =   -2147483634
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Periodo de Análisis"
         Begin SSCalendarWidgets_A.SSDateCombo dtFechaInicial 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   3
            EndProperty
            Height          =   345
            Left            =   330
            TabIndex        =   1
            Top             =   540
            Width           =   1995
            _Version        =   65537
            _ExtentX        =   3519
            _ExtentY        =   609
            _StockProps     =   93
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MinDate         =   "2008/1/1"
            MaxDate         =   "2015/12/31"
            BevelColorFace  =   -2147483634
            BevelColorFrame =   -2147483634
            ShowCentury     =   -1  'True
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtFechaFinal 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   3
            EndProperty
            Height          =   345
            Left            =   2490
            TabIndex        =   2
            Top             =   540
            Width           =   2025
            _Version        =   65537
            _ExtentX        =   3572
            _ExtentY        =   609
            _StockProps     =   93
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MinDate         =   "2008/1/1"
            MaxDate         =   "2015/12/31"
            BevelColorFace  =   -2147483634
            BevelColorFrame =   -2147483634
         End
         Begin VB.Label Label19 
            BackColor       =   &H8000000E&
            Caption         =   "Inicio:"
            Height          =   165
            Left            =   330
            TabIndex        =   8
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label20 
            BackColor       =   &H8000000E&
            Caption         =   "Fin:"
            Height          =   165
            Left            =   2490
            TabIndex        =   7
            Top             =   270
            Width           =   885
         End
      End
      Begin Threed.SSOption rbDineroCtesCobrados 
         Height          =   225
         Index           =   1
         Left            =   6180
         TabIndex        =   4
         Top             =   150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         _Version        =   196608
         BackColor       =   -2147483634
         Caption         =   "No. de Cobros"
      End
   End
   Begin Threed.SSPanel pnlGraficos 
      Height          =   5355
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   390
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   9446
      _Version        =   196608
      Caption         =   "1"
      RoundedCorners  =   0   'False
      Begin Threed.SSPanel pnlAnual 
         Height          =   4995
         Index           =   1
         Left            =   30
         TabIndex        =   13
         Top             =   330
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   8811
         _Version        =   196608
         RoundedCorners  =   0   'False
         Begin VtChartLib6.VtChart grResumenAnual 
            Height          =   4335
            Left            =   210
            TabIndex        =   14
            Top             =   180
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   7646
            _0              =   $"resumenGraficofrm.frx":15DDE
            _1              =   $"resumenGraficofrm.frx":161E7
            _2              =   $"resumenGraficofrm.frx":165F1
            _3              =   $"resumenGraficofrm.frx":169FA
            _4              =   $"resumenGraficofrm.frx":16E03
            _5              =   $"resumenGraficofrm.frx":1720C
            _6              =   $"resumenGraficofrm.frx":17615
            _7              =   $"resumenGraficofrm.frx":17A1E
            _8              =   $"resumenGraficofrm.frx":17E27
            _9              =   $"resumenGraficofrm.frx":18231
            _10             =   $"resumenGraficofrm.frx":1863A
            _11             =   $"resumenGraficofrm.frx":18A43
            _12             =   $"resumenGraficofrm.frx":18E4C
            _13             =   $"resumenGraficofrm.frx":19255
            _14             =   $"resumenGraficofrm.frx":1965E
            _15             =   $"resumenGraficofrm.frx":19A67
            _16             =   $"resumenGraficofrm.frx":19E70
            _17             =   $"resumenGraficofrm.frx":1A279
            _18             =   $"resumenGraficofrm.frx":1A682
            _19             =   $"resumenGraficofrm.frx":1AA8B
            _20             =   $"resumenGraficofrm.frx":1AE94
            _21             =   $"resumenGraficofrm.frx":1B29D
            _22             =   $"resumenGraficofrm.frx":1B6A6
            _23             =   $"resumenGraficofrm.frx":1BAAF
            _24             =   $"resumenGraficofrm.frx":1BEB8
            _25             =   $"resumenGraficofrm.frx":1C2C1
            _26             =   $"resumenGraficofrm.frx":1C6CA
            _27             =   $"resumenGraficofrm.frx":1CAD3
            _28             =   $"resumenGraficofrm.frx":1CEDC
            _29             =   $"resumenGraficofrm.frx":1D2E5
            _30             =   $"resumenGraficofrm.frx":1D6EE
            _31             =   $"resumenGraficofrm.frx":1DAF7
            _32             =   $"resumenGraficofrm.frx":1DF00
            _33             =   $"resumenGraficofrm.frx":1E309
            _34             =   $"resumenGraficofrm.frx":1E712
            _35             =   $"resumenGraficofrm.frx":1EB1B
            _36             =   $"resumenGraficofrm.frx":1EF24
            _37             =   $"resumenGraficofrm.frx":1F32D
            _38             =   $"resumenGraficofrm.frx":1F736
            _39             =   $"resumenGraficofrm.frx":1FB3F
            _40             =   $"resumenGraficofrm.frx":1FF48
            _41             =   $"resumenGraficofrm.frx":20351
            _42             =   $"resumenGraficofrm.frx":2075A
            _43             =   $"resumenGraficofrm.frx":20B63
            _44             =   $"resumenGraficofrm.frx":20F6C
            _45             =   $"resumenGraficofrm.frx":21375
            _46             =   $"resumenGraficofrm.frx":2177E
            _47             =   $"resumenGraficofrm.frx":21B87
            _48             =   $"resumenGraficofrm.frx":21F90
            _49             =   $"resumenGraficofrm.frx":22399
            _50             =   $"resumenGraficofrm.frx":227A2
            _51             =   $"resumenGraficofrm.frx":22BAB
            _52             =   $"resumenGraficofrm.frx":22FB4
            _53             =   $"resumenGraficofrm.frx":233BD
            _54             =   $"resumenGraficofrm.frx":237C6
            _55             =   $"resumenGraficofrm.frx":23BCF
            _56             =   $"resumenGraficofrm.frx":23FD8
            _57             =   $"resumenGraficofrm.frx":243E1
            _58             =   $"resumenGraficofrm.frx":247EA
            _59             =   $"resumenGraficofrm.frx":24BF3
            _60             =   $"resumenGraficofrm.frx":24FFC
            _61             =   $"resumenGraficofrm.frx":25405
            _62             =   $"resumenGraficofrm.frx":2580E
            _63             =   $"resumenGraficofrm.frx":25C17
            _64             =   $"resumenGraficofrm.frx":2601F
            _count          =   65
            _ver            =   2
         End
      End
      Begin Threed.SSPanel pnlAnual 
         Height          =   4995
         Index           =   2
         Left            =   30
         TabIndex        =   15
         Top             =   330
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   8811
         _Version        =   196608
         RoundedCorners  =   0   'False
         Begin EditLib.fpDoubleSingle fpdPorcentajeCrecimiento 
            Height          =   345
            Left            =   3930
            TabIndex        =   16
            Top             =   330
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            ControlType     =   1
            Text            =   "0"
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "99.99"
            MinValue        =   "0"
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
         Begin VtChartLib6.VtChart grCrecimiento 
            Height          =   4155
            Left            =   0
            TabIndex        =   17
            Top             =   690
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   7329
            _0              =   $"resumenGraficofrm.frx":26122
            _1              =   $"resumenGraficofrm.frx":2652B
            _2              =   $"resumenGraficofrm.frx":26935
            _3              =   $"resumenGraficofrm.frx":26D3E
            _4              =   $"resumenGraficofrm.frx":27147
            _5              =   $"resumenGraficofrm.frx":27550
            _6              =   $"resumenGraficofrm.frx":27959
            _7              =   $"resumenGraficofrm.frx":27D62
            _8              =   $"resumenGraficofrm.frx":2816B
            _9              =   $"resumenGraficofrm.frx":28574
            _10             =   $"resumenGraficofrm.frx":2897D
            _11             =   $"resumenGraficofrm.frx":28D86
            _12             =   $"resumenGraficofrm.frx":2918F
            _13             =   $"resumenGraficofrm.frx":29598
            _14             =   $"resumenGraficofrm.frx":299A2
            _15             =   $"resumenGraficofrm.frx":29DAB
            _16             =   $"resumenGraficofrm.frx":2A1B4
            _17             =   $"resumenGraficofrm.frx":2A5BD
            _18             =   $"resumenGraficofrm.frx":2A9C6
            _19             =   $"resumenGraficofrm.frx":2ADCF
            _20             =   $"resumenGraficofrm.frx":2B1D8
            _21             =   $"resumenGraficofrm.frx":2B5E1
            _22             =   $"resumenGraficofrm.frx":2B9EA
            _23             =   $"resumenGraficofrm.frx":2BDF3
            _24             =   $"resumenGraficofrm.frx":2C1FC
            _25             =   $"resumenGraficofrm.frx":2C605
            _26             =   $"resumenGraficofrm.frx":2CA0E
            _27             =   $"resumenGraficofrm.frx":2CE17
            _28             =   $"resumenGraficofrm.frx":2D220
            _29             =   $"resumenGraficofrm.frx":2D629
            _30             =   $"resumenGraficofrm.frx":2DA32
            _31             =   $"resumenGraficofrm.frx":2DE3B
            _32             =   $"resumenGraficofrm.frx":2E245
            _33             =   $"resumenGraficofrm.frx":2E64E
            _34             =   $"resumenGraficofrm.frx":2EA57
            _35             =   $"resumenGraficofrm.frx":2EE60
            _36             =   $"resumenGraficofrm.frx":2F269
            _37             =   $"resumenGraficofrm.frx":2F672
            _38             =   $"resumenGraficofrm.frx":2FA7B
            _39             =   $"resumenGraficofrm.frx":2FE84
            _40             =   $"resumenGraficofrm.frx":3028D
            _41             =   $"resumenGraficofrm.frx":30696
            _42             =   $"resumenGraficofrm.frx":30A9F
            _43             =   $"resumenGraficofrm.frx":30EA8
            _44             =   $"resumenGraficofrm.frx":312B1
            _45             =   $"resumenGraficofrm.frx":316BB
            _46             =   $"resumenGraficofrm.frx":31AC4
            _47             =   $"resumenGraficofrm.frx":31ECD
            _48             =   $"resumenGraficofrm.frx":322D7
            _49             =   $"resumenGraficofrm.frx":326E0
            _50             =   $"resumenGraficofrm.frx":32AE9
            _51             =   $"resumenGraficofrm.frx":32EF2
            _52             =   $"resumenGraficofrm.frx":332FB
            _53             =   $"resumenGraficofrm.frx":33704
            _54             =   $"resumenGraficofrm.frx":33B0D
            _55             =   $"resumenGraficofrm.frx":33F16
            _56             =   $"resumenGraficofrm.frx":34320
            _57             =   $"resumenGraficofrm.frx":34729
            _58             =   $"resumenGraficofrm.frx":34B32
            _59             =   $"resumenGraficofrm.frx":34F3C
            _60             =   $"resumenGraficofrm.frx":35345
            _61             =   $"resumenGraficofrm.frx":3574E
            _62             =   $"resumenGraficofrm.frx":35B57
            _63             =   $"resumenGraficofrm.frx":35F60
            _64             =   $"resumenGraficofrm.frx":36369
            _65             =   $"resumenGraficofrm.frx":36772
            _66             =   $"resumenGraficofrm.frx":36B7B
            _67             =   $"resumenGraficofrm.frx":36F84
            _68             =   $"resumenGraficofrm.frx":3738D
            _69             =   $"resumenGraficofrm.frx":37796
            _70             =   $"resumenGraficofrm.frx":37B9F
            _71             =   $"resumenGraficofrm.frx":37FA8
            _72             =   $"resumenGraficofrm.frx":383B1
            _73             =   $"resumenGraficofrm.frx":387BA
            _74             =   $"resumenGraficofrm.frx":38BC3
            _75             =   $"resumenGraficofrm.frx":38FCC
            _76             =   $"resumenGraficofrm.frx":393D5
            _77             =   $"resumenGraficofrm.frx":397DE
            _78             =   $"resumenGraficofrm.frx":39BE7
            _79             =   $"resumenGraficofrm.frx":39FF0
            _80             =   $"resumenGraficofrm.frx":3A3F9
            _81             =   $"resumenGraficofrm.frx":3A802
            _82             =   $"resumenGraficofrm.frx":3AC0B
            _83             =   $"resumenGraficofrm.frx":3B014
            _84             =   $"resumenGraficofrm.frx":3B41D
            _85             =   $"resumenGraficofrm.frx":3B826
            _86             =   $"resumenGraficofrm.frx":3BC2F
            _87             =   $"resumenGraficofrm.frx":3C038
            _88             =   $"resumenGraficofrm.frx":3C441
            _89             =   $"resumenGraficofrm.frx":3C84A
            _90             =   $"resumenGraficofrm.frx":3CC53
            _91             =   $"resumenGraficofrm.frx":3D05C
            _92             =   $"resumenGraficofrm.frx":3D465
            _93             =   $"resumenGraficofrm.frx":3D86E
            _94             =   "-@@D@@@@@A@@@@@.????*@*@????????.A@@@@@e@@@I@)V4-Window-@C@D@@@@@,5356"
            _count          =   95
            _ver            =   2
         End
         Begin VB.Label Label1 
            Caption         =   "% Crecimiento Total:"
            Height          =   285
            Left            =   2400
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
      End
      Begin Threed.SSPanel pnlAnual 
         Height          =   4995
         Index           =   3
         Left            =   30
         TabIndex        =   19
         Top             =   330
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   8811
         _Version        =   196608
         RoundedCorners  =   0   'False
         Begin VtChartLib6.VtChart grPrestamoCobranza 
            Height          =   4635
            Left            =   210
            TabIndex        =   20
            Top             =   180
            Width           =   8385
            _ExtentX        =   14790
            _ExtentY        =   8176
            _0              =   $"resumenGraficofrm.frx":3DC77
            _1              =   $"resumenGraficofrm.frx":3E080
            _2              =   $"resumenGraficofrm.frx":3E489
            _3              =   $"resumenGraficofrm.frx":3E892
            _4              =   $"resumenGraficofrm.frx":3EC9B
            _5              =   $"resumenGraficofrm.frx":3F0A4
            _6              =   $"resumenGraficofrm.frx":3F4AD
            _7              =   $"resumenGraficofrm.frx":3F8B6
            _8              =   $"resumenGraficofrm.frx":3FCBF
            _9              =   $"resumenGraficofrm.frx":400C8
            _10             =   $"resumenGraficofrm.frx":404D1
            _11             =   $"resumenGraficofrm.frx":408DA
            _12             =   $"resumenGraficofrm.frx":40CE3
            _13             =   $"resumenGraficofrm.frx":410EC
            _14             =   $"resumenGraficofrm.frx":414F5
            _15             =   $"resumenGraficofrm.frx":418FE
            _16             =   $"resumenGraficofrm.frx":41D08
            _17             =   $"resumenGraficofrm.frx":42111
            _18             =   $"resumenGraficofrm.frx":4251A
            _19             =   $"resumenGraficofrm.frx":42923
            _20             =   $"resumenGraficofrm.frx":42D2C
            _21             =   $"resumenGraficofrm.frx":43136
            _22             =   $"resumenGraficofrm.frx":4353F
            _23             =   $"resumenGraficofrm.frx":43948
            _24             =   $"resumenGraficofrm.frx":43D51
            _25             =   $"resumenGraficofrm.frx":4415A
            _26             =   $"resumenGraficofrm.frx":44563
            _27             =   $"resumenGraficofrm.frx":4496C
            _28             =   $"resumenGraficofrm.frx":44D76
            _count          =   29
            _ver            =   2
         End
      End
      Begin ComctlLib.TabStrip tabAnual 
         Height          =   5355
         Left            =   30
         TabIndex        =   12
         Top             =   0
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9446
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   3
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Comparativo"
               Key             =   "COMPARATIVO"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Crecimiento"
               Key             =   "CRECIMIENTO"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Prestamo vs Cobranza"
               Key             =   "PRESTAMO_COBRANZA"
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip tabGrafico 
      Height          =   5715
      Left            =   30
      TabIndex        =   10
      Top             =   60
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   10081
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Porcentajes"
            Key             =   "PORCENTAJES"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Anual"
            Key             =   "ANUAL"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   660
      Top             =   5070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196608
      MinFontSize     =   1
      MaxFontSize     =   100
   End
End
Attribute VB_Name = "resumenGraficofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miFrameActivoOP As Integer
Private miFrameActivoAnual As Integer
Private miFrameActivoMensual As Integer


Private Const TAB_PORCENTAJES = 0
Private Const TAB_PROMEDIO_MENSUAL = 4
Private Const TAB_RESUMEN_ANUAL = 1

Private Const TAB_COMPARATIVO = 1
Private Const TAB_CRECIMIENTO = 2
Private Const TAB_PRESTADO_COBRANZA = 3

Private Const TAB_PROMEDIO = 1
Private Const TAB_ESTADO_RESULTADOS = 2

Private Const COL_ATRASO = 7

Private Const PERIODO_ACUMULADO = 0
Private Const PERIODO_DIARIO = 1
Private Const PERIODO_SEMANAL = 2
Private Const PERIODO_MENSUAL = 3
Private Const PERIODO_TRIMESTRAL = 4
Private Const PERIODO_ANUAL = 5
Private Const COBRO_MONTO = 1
Private Const COBRO_NUMERO = 0

Private strCabecera As String

Private bConsultaHecha As Boolean
Private bConsultaCrecimientoHecha As Boolean
Private bConsultaPrestadoCobradoHecha As Boolean
Private bCatalogCobradoresCargado As Boolean


'Private Sub cmdAceptar_Click()
'
'    Dim oPago As New Pago
'
'    oPago.analisisCredito cbCobrador
'
'    If oPago.cDatos.Count > 0 Then
'        Call fnLlenaTablaCollection(sprAnalisis, oPago.cDatos)
'
'        txtMontoAtrasado.Text = Format(obtenTotalGrid(sprAnalisis, COL_ATRASO), "###,###,###,###0.00")
'
'    End If
'
'    bCatalogCobradoresCargado = True
'
'    Set oPago = Nothing
'
'End Sub

Private Sub Form_Load()
    
    miFrameActivoOP = TAB_PORCENTAJES
    miFrameActivoAnual = TAB_COMPARATIVO
    miFrameActivoMensual = TAB_PROMEDIO
    strCabecera = sicPrincipalfrm.pnlTitulo.Caption
    
    bConsultaHecha = False
    bConsultaCrecimientoHecha = False
    bConsultaPrestadoCobradoHecha = False
    bCatalogCobradoresCargado = False
    
    dtFechaInicial = Date
    dtFechaFinal = Date
    cbPeriodo.ListIndex = PERIODO_DIARIO
    
    generaGrafica
    
End Sub

Private Sub generaGrafica()
    
    Dim oPago As New Pago

    If rbDineroCtesCobrados(0).Value = True Then 'Dinero cobrado
        
        Select Case cbPeriodo.Text
            Case Is = "Acumulado"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, dtFechaFinal.Text, PERIODO_ACUMULADO, COBRO_MONTO, 1), _
                                                               grResumen, _
                                                               0)
                sicPrincipalfrm.pnlTitulo.Caption = cbPeriodo.Text & " - Porcentaje Monto Cobrado"

            Case Is = "Diario"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                             dtFechaFinal.Text, PERIODO_DIARIO, COBRO_MONTO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje Monto Cobrado"

            Case Is = "Semanal"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                             dtFechaFinal.Text, PERIODO_SEMANAL, COBRO_MONTO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje Monto Cobrado"

            Case Is = "Mensual"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                             dtFechaFinal.Text, PERIODO_MENSUAL, COBRO_MONTO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje Monto Cobrado"

            Case Is = "Trimestral"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                             dtFechaFinal.Text, PERIODO_TRIMESTRAL, COBRO_MONTO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje Monto Cobrado"

            Case Is = "Anual"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                             dtFechaFinal.Text, PERIODO_ANUAL, COBRO_MONTO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje Monto Cobrado"

        End Select
    Else                'No de cobros
        Select Case cbPeriodo.Text
            Case Is = "Acumulado"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                          dtFechaFinal.Text, PERIODO_ACUMULADO, COBRO_NUMERO, 1), grResumen, 0)
                sicPrincipalfrm.pnlTitulo.Caption = cbPeriodo.Text & " - Porcentaje No de Cobros"

            Case Is = "Diario"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                          dtFechaFinal.Text, PERIODO_DIARIO, COBRO_NUMERO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje No de Cobros"

            Case Is = "Semanal"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                          dtFechaFinal.Text, PERIODO_SEMANAL, COBRO_NUMERO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje No de Cobros"

            Case Is = "Mensual"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                          dtFechaFinal.Text, PERIODO_MENSUAL, COBRO_NUMERO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje No de Cobros"

            Case Is = "Trimestral"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                          dtFechaFinal.Text, PERIODO_TRIMESTRAL, COBRO_NUMERO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje No de Cobros"

            Case Is = "Anual"
                Call dibujaGrafica(oPago.acumuladoPorcentajes(dtFechaInicial.Text, _
                                                          dtFechaFinal.Text, PERIODO_ANUAL, COBRO_NUMERO, 1), grResumen, 1)
                sicPrincipalfrm.pnlTitulo.Caption = "Periodo " & cbPeriodo.Text & " - Porcentaje No de Cobros"

        End Select

    End If

    Set oPago = Nothing

End Sub

Private Sub dtFechaInicial_Change()
   
   If validaCaptura() = True Then
    
        generaGrafica
        
    End If

End Sub

Private Sub dtFechaInicial_Click()
    
    If validaCaptura() = True Then
    
        generaGrafica
        
    End If

End Sub

Private Sub dtFechaInicial_CloseUp()
    
    If validaCaptura() = True Then
    
        generaGrafica
    End If

End Sub

Private Sub dtFechaFinal_Change()
   
   If validaCaptura() = True Then
    
        generaGrafica
        
    End If

End Sub

Private Sub dtFechaFinal_Click()
    
    If validaCaptura() = True Then
    
        generaGrafica
        
    End If

End Sub

Private Sub dtFechaFinal_CloseUp()
    
    If validaCaptura() = True Then
    
        generaGrafica
        
    End If

End Sub

Private Sub cbPeriodo_Click()

    If validaCaptura() = True Then
    
        generaGrafica
        
    End If
    
End Sub

Private Sub rbDineroCtesCobrados_Click(Index As Integer, Value As Integer)

    If validaCaptura() = True Then
    
        generaGrafica
        
    End If
    
End Sub

Private Function validaCaptura() As Boolean

    validaCaptura = False
    If Len(dtFechaInicial) = 0 Then
        dtFechaInicial.SetFocus
        Exit Function
    End If
    
    If DateDiff("d", dtFechaInicial, dtFechaFinal) < 0 Then
        MsgBox "Verifique el perio de fechas, la inicial debe ser menor o igual a la final.", vbOKOnly
        dtFechaInicial.SetFocus
        Exit Function
    End If

    If rbDineroCtesCobrados(0).Value = False And rbDineroCtesCobrados(1) = False Then
        MsgBox "Selecciona por Montos o por No. de Cobros", vbOKOnly
        rbDineroCtesCobrados(0).SetFocus
        Exit Function
    End If
    
'    If cbPeriodo.ListIndex = -1 Then
'
'        MsgBox "Defina periodo para el reporte", vbOKOnly
'        cbPeriodo.SetFocus
'        Exit Function
'
'    End If
    
    validaCaptura = True

End Function

Private Sub tabAnual_Click()

    Dim oPago As New Pago

    If tabAnual.SelectedItem.Index = miFrameActivoAnual Then Exit Sub ' No need to change frame.
    
    ' Comosea, oculta el frame anterior, muestra el nuevo.
    pnlAnual(tabAnual.SelectedItem.Index).Visible = True
    pnlAnual(miFrameActivoAnual).Visible = False
    
    miFrameActivoAnual = tabAnual.SelectedItem.Index
    
    Select Case miFrameActivoAnual
        Case Is = TAB_COMPARATIVO
            strCabecera = sicPrincipalfrm.pnlTitulo.Caption
            sicPrincipalfrm.pnlTitulo.Caption = "Comparativo anual de Cobranza"
            
            If bConsultaHecha = False Then
                
                Call dibujaGrafica(oPago.acumuladoAnual, grResumenAnual, 1)
                            
            End If
        Case Is = TAB_CRECIMIENTO
            
            If bConsultaCrecimientoHecha = False Then
                
                Dim cRegistros As New Collection
                Dim cRegistro As New Collection
                Dim oCampo As New Campo
                
                strCabecera = sicPrincipalfrm.pnlTitulo.Caption
                sicPrincipalfrm.pnlTitulo.Caption = "Comparativo anual de Crecimiento"
                
                Call dibujaGrafica(oPago.crecimientoAnual, grCrecimiento, 0)
                            
                Set cRegistros = oPago.crecimientoTotal
                Set cRegistro = cRegistros(1)
                Set oCampo = cRegistro(1)
                
                If IsNull(oCampo.Valor) Then
                    fpdPorcentajeCrecimiento = "0"
                Else
                    fpdPorcentajeCrecimiento.Text = oCampo.Valor
                End If
            End If
        Case Is = TAB_PRESTADO_COBRANZA

            If bConsultaPrestadoCobradoHecha = False Then
                
                strCabecera = sicPrincipalfrm.pnlTitulo.Caption
                sicPrincipalfrm.pnlTitulo.Caption = "Comparativo anual de Préstado .vs. Cobrado"
                Call dibujaGrafica(oPago.comparativoPrestamoCobranza, grPrestamoCobranza, 1)
            
            End If
            
        Case Else
            sicPrincipalfrm.pnlTitulo.Caption = strCabecera

    End Select
    
    Set oPago = Nothing

End Sub


Private Function obtenFrame(key As String) As Integer
    
    Dim iFrame As Integer
    Select Case key
        Case Is = "PORCENTAJES"
            iFrame = 0
        Case Is = "ANUAL"
            iFrame = 1
    End Select

    obtenFrame = iFrame
    
End Function

Private Sub tabGrafico_Click()
    
    pnlGraficos(obtenFrame(tabGrafico.SelectedItem.key)).ZOrder 0

    'If tabGrafico.SelectedItem.Index = miFrameActivoOP Then Exit Sub ' No need to change frame.
    
    ' Comosea, oculta el frame anterior, muestra el nuevo.
    'pnlGraficos(tabGrafico.SelectedItem.Index).Visible = True
    'pnlGraficos(miFrameActivoOP).Visible = False
    
    miFrameActivoOP = tabGrafico.SelectedItem.Index
    
    Select Case miFrameActivoOP
        
        Case Is = TAB_PORCENTAJES
        
        Case Is = TAB_PROMEDIO_MENSUAL
            
'            If bCatalogCobradoresCargado = False Then
'
'                'Carga el catalogo de cobradores
'                Dim oUsuario As New Usuario
'                If oUsuario.catalogoUsuarios Then
'                    fnLlenaComboCollecion cbCobrador, oUsuario.cDatos, 0, ""
'                End If
'                Set oUsuario = Nothing
'
'            End If
            
        Case Is = TAB_RESUMEN_ANUAL
        
            If miFrameActivoAnual = TAB_COMPARATIVO Then
                strCabecera = sicPrincipalfrm.pnlTitulo.Caption
                sicPrincipalfrm.pnlTitulo.Caption = "Comparativo de Cobranza"
        
                If bConsultaHecha = False Then
                    Dim oPago As New Pago
                    Call dibujaGrafica(oPago.acumuladoAnual, grResumenAnual, 1)
                    Set oPago = Nothing
                End If
        
            Else
                sicPrincipalfrm.pnlTitulo.Caption = strCabecera
            End If
    End Select
    
    
End Sub

'Private Sub tabMensual_Click()
'
'    If tabMensual.SelectedItem.Index = miFrameActivoMensual Then Exit Sub ' No need to change frame.
'
'    ' Comosea, oculta el frame anterior, muestra el nuevo.
'    pnlMensual(tabMensual.SelectedItem.Index).Visible = True
'    pnlMensual(miFrameActivoMensual).Visible = False
'
'    miFrameActivoMensual = tabMensual.SelectedItem.Index
'
'End Sub


