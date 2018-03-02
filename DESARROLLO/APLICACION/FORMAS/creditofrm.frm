VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form creditofrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito del Cliente"
   ClientHeight    =   6240
   ClientLeft      =   1860
   ClientTop       =   2580
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtnocliente 
      Height          =   330
      Left            =   1230
      TabIndex        =   38
      Top             =   225
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####"
      PromptChar      =   "_"
   End
   Begin EditLib.fpCurrency txtSaldoCuenta 
      Height          =   345
      Left            =   1500
      TabIndex        =   36
      Top             =   5610
      Visible         =   0   'False
      Width           =   2025
      _Version        =   196608
      _ExtentX        =   3572
      _ExtentY        =   609
      Enabled         =   0   'False
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
   Begin VB.ComboBox cbCuentaCheques 
      Height          =   360
      Left            =   1560
      TabIndex        =   33
      Top             =   5160
      Visible         =   0   'False
      Width           =   2025
   End
   Begin Crystal.CrystalReport crFactura 
      Left            =   4260
      Top             =   5130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   2385
      Left            =   -30
      TabIndex        =   13
      Top             =   1110
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4207
      _Version        =   196608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Generales del Crédito"
      Begin MSMask.MaskEdBox txtfactura 
         Height          =   330
         Left            =   1245
         TabIndex        =   45
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "######"
         PromptChar      =   "_"
      End
      Begin SSCalendarWidgets_A.SSDateCombo dtFechaRegistro 
         Height          =   345
         Left            =   5880
         TabIndex        =   29
         Top             =   330
         Width           =   1965
         _Version        =   65537
         _ExtentX        =   3466
         _ExtentY        =   609
         _StockProps     =   93
         Enabled         =   0   'False
         ScrollBarTracking=   0   'False
         SpinButton      =   0
      End
      Begin Threed.SSFrame SSFrame3 
         Height          =   975
         Left            =   180
         TabIndex        =   24
         Top             =   1230
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   1720
         _Version        =   196608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Periodo de Cobranza"
         Begin SSCalendarWidgets_A.SSDateCombo dtFechaCobranzaInicial 
            Height          =   345
            Left            =   330
            TabIndex        =   25
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
         End
         Begin SSCalendarWidgets_A.SSDateCombo dtFechaCobranzaFinal 
            Height          =   345
            Left            =   2490
            TabIndex        =   26
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
         End
         Begin VB.Label Label20 
            Caption         =   "Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   2490
            TabIndex        =   28
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label19 
            Caption         =   "Inicio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   330
            TabIndex        =   27
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.TextBox txtstatus 
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
         Left            =   5880
         TabIndex        =   16
         Text            =   "V"
         Top             =   1230
         Width           =   375
      End
      Begin VB.TextBox txtdescrip 
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
         Left            =   1245
         TabIndex        =   15
         Top             =   750
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.CheckBox opcion 
         Alignment       =   1  'Right Justify
         Caption         =   "Electricos"
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
         Left            =   4890
         TabIndex        =   14
         Top             =   795
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Folio:"
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
         Left            =   210
         TabIndex        =   23
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Status:"
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
         Left            =   4920
         TabIndex        =   21
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "V- Vigente"
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
         Left            =   5970
         TabIndex        =   20
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "T-Terminado"
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
         Left            =   6840
         TabIndex        =   19
         Top             =   1950
         Width           =   945
      End
      Begin VB.Label Label15 
         Caption         =   "P-Pendiente"
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
         Left            =   4920
         TabIndex        =   18
         Top             =   1950
         Width           =   975
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción:"
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
         Left            =   210
         TabIndex        =   17
         Top             =   750
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1515
      Left            =   -30
      TabIndex        =   5
      Top             =   3510
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   2672
      _Version        =   196608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Definición del Crédito:"
      Begin MSMask.MaskEdBox txtfinan 
         Height          =   330
         Left            =   6525
         TabIndex        =   44
         Top             =   1080
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "$ #####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txttotalpagar 
         Height          =   330
         Left            =   6525
         TabIndex        =   43
         Top             =   675
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "$ ######"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPagos 
         Height          =   330
         Left            =   2340
         TabIndex        =   42
         Top             =   1035
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPorcentajeFinanciamiento 
         Height          =   330
         Left            =   2340
         TabIndex        =   41
         Top             =   675
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtpagodiario 
         Height          =   330
         Left            =   6525
         TabIndex        =   40
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "$ #####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCredito 
         Height          =   330
         Left            =   2340
         TabIndex        =   39
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "$ ######"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Monto del crédito:"
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
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Financiamiento:"
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
         Left            =   4410
         TabIndex        =   11
         Top             =   1050
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "No. de pagos:"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Cantidad a pagar:"
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
         Left            =   4410
         TabIndex        =   9
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad total a pagar:"
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
         Left            =   4410
         TabIndex        =   8
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Financiamiento en %:"
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
         Left            =   150
         TabIndex        =   7
         Top             =   690
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "%"
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
         Left            =   3690
         TabIndex        =   6
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCredito 
      Caption         =   "Grabar"
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
      Left            =   5055
      TabIndex        =   1
      Top             =   5670
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
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
      Left            =   6495
      TabIndex        =   2
      Top             =   5670
      Width           =   1335
   End
   Begin VB.TextBox txtnombre 
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
      Left            =   1230
      TabIndex        =   3
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label lblCobrador 
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
      Left            =   5700
      TabIndex        =   37
      Top             =   5160
      Width           =   2085
   End
   Begin VB.Label Label22 
      Caption         =   "Saldo:"
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Cuenta de Cheques"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   5160
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label9 
      Caption         =   "Cobrador:"
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
      Left            =   4950
      TabIndex        =   32
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label lblCheque 
      Alignment       =   1  'Right Justify
      Caption         =   "Cheque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1680
      TabIndex        =   31
      Top             =   6030
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label Label10 
      Caption         =   "Cheque:"
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   6030
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "No. de Cliente:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblCliente 
      Caption         =   "Cliente:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "creditofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iNoCliente As Integer
Public iPagos As Integer
Public iNoFactura As Long
Public strNombreCliente As String
Public bAltaCredito As Boolean

Private Const POS_NUM_CLIENTE = 1
Private Const POS_NUM_FACTURA = 2
Private Const POS_MONTO_CREDITO = 3
Private Const POS_CANTIDAD_PAGAR = 4
Private Const POS_FINANCIAMIENTO = 5
Private Const POS_CANTIDAD_TOTAL_PAGAR = 6
Private Const POS_NUMERO_PAGOS = 7
Private Const POS_FECHA_INICIO_CREDITO = 8
Private Const POS_FECHA_FINAL_CREDITO = 9
Private Const POS_FECHA_ALTA_CREDITO = 10
Private Const POS_STATUS_CREDITO = 11
Private Const POS_TIPO_CREDITO = 12
Private Const POS_DESC_CREDITO = 13

'Private Const POS_NOMBRE_CLINETE = 2
'Private Const POS_PORCENTAJE_FINANCIAMIENTO = 11

Private fFinanciamiento As Double
Private iNoPagos As Integer
Private fCantidadPagar As Double
Private chStatus As String
Private iTipoCredito As Integer

Private iVeces As Integer

Public strCobrador As String

Private bInicio As Boolean

Private Sub Form_Activate()
        
        'txtfactura.Text = ""
        txtfactura.Enabled = False
        txtcredito.Text = ""
        txtfinan.Text = 0
        txtPorcentajeFinanciamiento.Text = 14
        txtpagos.Text = 30
        txtpagodiario.Text = ""
        txttotalpagar.Text = ""
        dtFechaRegistro.Text = Format(Now(), "dd/mm/yyyy")
        dtFechaCobranzaInicial.Text = Format(Now() + 2, "dd/mm/yyyy")
        dtFechaCobranzaFinal.Text = Format(Now + CInt(txtpagos.Text) + 2, "dd/mm/yyyy")
        cmdcredito.Enabled = True
        txtstatus.Enabled = True
        lblCobrador.Caption = strCobrador

        txtcredito.SetFocus
        
        bInicio = False

End Sub

Private Sub Form_Load()

    Dim oCredito As New credito
    'Si el No. de factura es  = 0 indica que es alta (crèdito nuevo)
    If iNoFactura = 0 Then
            bInicio = True

        'Inicializa la forma, para captura de información
'        txtfactura.Text = ""
'        txtfactura.Enabled = False
'        txtCredito.Text = ""
'        txtfinan.Text = 0
'        txtPorcentajeFinanciamiento.Text = 14
'        txtPagos.Text = 30
'        txtpagodiario.Text = ""
'        txttotalpagar.Text = ""
'        dtFechaRegistro.Text = Format(Now(), "dd/mm/yyyy")
'        dtFechaCobranzaInicial.Text = Format(Now() + 2, "dd/mm/yyyy")
'        dtFechaCobranzaFinal.Text = Format(Now + CInt(txtPagos.Text) + 2, "dd/mm/yyyy")
'        cmdCredito.Enabled = True
'        txtstatus.Enabled = True

        'Al botón cmdCredito poner el caption 'Registrar'
        cmdcredito.Caption = "Registrar"
        
        'Obten el siguiente folio (factura)
        txtfactura.Text = oCredito.siguiente
        Set oCredito = Nothing
        
'        'Carga el catalogo de cobradores
'        Dim oUsuario As New Usuario
'        If oUsuario.catalogoUsuarios Then
'            fnLlenaComboCollecion cbCobrador, oUsuario.cDatos, 0, ""
'            'busca el cobrador en la lista y hazlo activo
'            fnBuscaTextoCombo cbCobrador, strCobrador
'        End If
'        Set oUsuario = Nothing
        
        'Dim strCantidad As String
        'strCantidad = convierteMontoConLentra("1000000.10")
        
'        'Carga catálogo CUENTAS DE CHEQUES
'        Dim oCtaCheques As New CuentaCheques
'
'        If oCtaCheques.catalogoEsp = True Then
'            Call fnLlenaComboCollecion(cbCuentaCheques, oCtaCheques.cDatos, 0, "")
'        Else
'
'            MsgBox "¡Es importante registrar una cuenta de cheques, para el control de sus finanzas!", vbInformation + vbOKOnly
'
'            cuentaChequesfrm.Show vbModal
'
'            If oCtaCheques.catalogoEsp = True Then
'                Call fnLlenaComboCollecion(cbCuentaCheques, oCtaCheques.cDatos, 0, "")
'            Else
'                MsgBox "¡Para el control de sus finanzas, registre por lo menos una cuenta de cheques!", vbCritical + vbOKOnly
'                Exit Sub
'            End If
'
'        End If
'
'        Set oCtaCheques = Nothing
        lblCheque.Caption = ""
        
    Else
    'Si el No. de factura es  <> de 0 indica que es cambio o consulta al crédito
        'Buscar los datos del credito, enviando el No. de cliente y el No. de Factura
            'Desplegar los datos del crédito (el desplegado se debe hacer en la función despliegaDatos)
            'Call despliegaDatos(oCredito.datosCredito(iNoCliente, iNoFactura))
            oCredito.datosCredito (iNoFactura)
            Call despliegaDatos(oCredito.cDatos)
            'Mantener los datos factibles de actualizar (Financiamiento, No. de pagos, cantidad a pagar, status y el tipo de crédito)
            ' en variables para su comparación y validación posterior.
            
            'Al botón cmdCredito poner el caption 'Actualizar'
            cmdcredito.Caption = "Actualizar"
        
        Set oCredito = Nothing
        
        txtnocliente.Enabled = False
        txtNombre.Enabled = False
        txtfactura.Enabled = False
        txtdescrip.Enabled = False
        dtFechaRegistro.Enabled = False
        
    End If
    txtnocliente = iNoCliente
    txtNombre = strNombreCliente
    
    'txtfactura.Enabled = False
    txtpagodiario.Enabled = False
    txttotalpagar.Enabled = False
    txtfinan.Enabled = False
    
    iVeces = 1
'COnsideraciones

    'Utilizar la clase Credito, para obtener la información de un credito segun cliente y factura dados
    'Utilizar la función datosCredito de la clase Credito
    
End Sub

Private Function despliegaDatos(cCredito As Collection)

    'Esta función recibe los datos de un crédito, en una collección.
    'Esta collección trae un registro tambien de tipo collección
    'el registro trae un conjunto de objetos de tipo campo (ver clase Campo)
    'Cada campo representa un dato del crédito, el orden en que viene cada campo en el registro es importante considerarlo
    'desde la consulta a la base de datos, para saber en que posición viene cada dato.
    
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    
    Set cRegistro = cCredito(1)
    
    Set oCampo = cRegistro(POS_NUM_CLIENTE)
    txtnocliente.Text = oCampo.Valor
    'Set oCampo = cRegistro(POS_NOMBRE_CLINETE)
    'txtnombre.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_NUM_FACTURA)
    txtfactura.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_FECHA_ALTA_CREDITO)
    dtFechaRegistro.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_DESC_CREDITO)
    If IsNull(oCampo.Valor) Then
        txtdescrip.Text = ""
    Else
        txtdescrip.Text = oCampo.Valor
    End If
    Set oCampo = cRegistro(POS_TIPO_CREDITO)
    If oCampo.Valor = 1 Then
        opcion.Value = True
    Else
        opcion.Value = False
    End If
    Set oCampo = cRegistro(POS_STATUS_CREDITO)
    txtstatus.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_FECHA_INICIO_CREDITO)
    dtFechaCobranzaInicial.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_FECHA_FINAL_CREDITO)
    dtFechaCobranzaFinal.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_MONTO_CREDITO)
    'txtCredito.Text = oCampo.Valor
    'Set oCampo = cRegistro(POS_PORCENTAJE_FINANCIAMIENTO)
    txtPorcentajeFinanciamiento.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_NUMERO_PAGOS)
    txtpagos.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_CANTIDAD_PAGAR)
    txtpagodiario.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_CANTIDAD_TOTAL_PAGAR)
    txttotalpagar.Text = oCampo.Valor
    Set oCampo = cRegistro(POS_FINANCIAMIENTO)
    txtfinan.Text = oCampo.Valor
    
'Consideraciones
    'Cada posición (indice del conjunto de datos) NO debe estar indicada por un número, sino que debe ser una constante definida
    'en la zona de declaraciones de este código.
    
End Function

Private Function validaForma() As Boolean
    
    Dim strMensaje As String
    Dim iFechaControl As Integer
    
    validaForma = False
    'Esta función valida cada uno de los datos que se capturan en la forma
    'En caso de que no haya información capturada y esta sea importante, debe enviar un mensaje
    'solicitando al usuario, capture el dato y debe colocar el focus en el control correspondiente.
    If txtnocliente.Enabled = True Then
        If txtnocliente.Text = "" Then
            MsgBox "Por favor defina el No. de Cliente.", vbInformation + vbOKOnly
            txtnocliente.SetFocus
            Exit Function
        End If
    End If
    
    If txtNombre.Text = "" Then
        MsgBox "No ha definido el nombre del cliente, Por favor defina este.", vbInformation + vbOKOnly
        txtNombre.SetFocus
        Exit Function
    End If
    
'    If txtfactura.Text = "" Then
'        MsgBox "No ha definido el No. de Folio, Por favor defina este.", vbInformation + vbOKOnly
'        txtfactura.SetFocus
'        Exit Function
'    End If
'
'    If Val(txtfactura.Text) <= 0 Then
'        MsgBox "No ha definido el No. de Folio, Por favor defina este.", vbInformation + vbOKOnly
'        txtfactura.SetFocus
'        Exit Function
'    End If
    
    'dtFechaRegistro.Text
    
    If opcion.Value = True Then
        If txtdescrip.Text = "" Then
            If vbNo = MsgBox("¿Es correcto que NO haya descripción para el folio?", vbQuestion + vbYesNo) Then
                txtdescrip.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If txtstatus.Text = "" Then
        MsgBox "No ha definido el estatus del Cédito, defina este por favor!", vbInformation + vbOKOnly
        txtstatus.SetFocus
        Exit Function
    End If
    
    strMensaje = fnValidaPeriodoFechas(dtFechaCobranzaInicial.Text, dtFechaCobranzaFinal.Text, iFechaControl)
    
    If strMensaje <> "" Then
        MsgBox strMensaje, vbInformation + vbOKOnly
        If iFechaControl = 1 Then
            dtFechaCobranzaInicial.SetFocus
        End If
        If iFechaControl = 2 Then
            dtFechaCobranzaFinal.SetFocus
        End If
        Exit Function
    End If
    
    If Val(fnstrValor(txtcredito.Text)) <= 0 Then
        MsgBox "Por favor defina el monto de crédito!", vbInformation + vbOKOnly
        txtcredito.SetFocus
        Exit Function
    End If
    
    If Val(fnstrValor(txtPorcentajeFinanciamiento.Text)) <= 0 Then
        MsgBox "Por favor defina el porcentaje de financiamiento!", vbInformation + vbOKOnly
        txtPorcentajeFinanciamiento.SetFocus
        Exit Function
    End If
    
'    If cbCuentaCheques.Text = "" Then
'        MsgBox "Por favor seleccione una cuenta de cheques!", vbInformation + vbOKOnly
'        cbCuentaCheques.SetFocus
'        Exit Function
'    End If
    
    'If cbCobrador.Text = "" Then
    '    MsgBox "Por favor seleccione quien cobrará este crédito!", vbInformation + vbOKOnly
    '    cbCobrador.SetFocus
    '    Exit Function
    'End If
    
'    If Val(fnstrValor(txtSaldoCuenta.Text)) < Val(fnstrValor(txtcredito.Text)) Then
'        MsgBox "No hay suficiente saldo en la cuenta, verifique porfavor!", vbInformation + vbOKOnly
'        cbCuentaCheques.SetFocus
'        Exit Function
'    End If
    
    'txtPagos.Text
    'txtpagodiario.Text
    'txttotalpagar.Text
    'txtfinan.Text
    
    validaForma = True
    
End Function

Private Function obtenDatos() As Collection

    'Esta función recoge cada dato y lo ingresa en una estructura de control 'Campo'
    'Cada campo lo agrega a una estrucutra 'collection' para armar un registro
    'El registro con los campos se agrega a otra estructura 'collection', la cual contiene registros (en este caso es solo uno)
    'la función regresa la colección de registros (en este caso es solo uno)
    Dim cCredito As New Collection
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    
    If cmdcredito.Caption = "Registrar" Then
    
        cRegistro.Add oCampo.CreaCampo(adInteger, , , txtnocliente.Text)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtcredito.Text)))   'Monto otorgado en el crèdito
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtpagodiario.Text))) 'Cantidad a pagar diario
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtfinan.Text))) 'Total a pagar de financiamiento
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txttotalpagar.Text))) 'Cantidad total a pagar
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtpagos.Text)) 'Numero de pagos
        cRegistro.Add oCampo.CreaCampo(adInteger, , , dtFechaCobranzaInicial.Text)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , dtFechaCobranzaFinal.Text)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , dtFechaRegistro.Text)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , txtstatus.Text)
        If opcion.Value = 1 Then
            cRegistro.Add oCampo.CreaCampo(adInteger, , , 1)
            cRegistro.Add oCampo.CreaCampo(adInteger, , , txtdescrip.Text)
        Else
            cRegistro.Add oCampo.CreaCampo(adInteger, , , 0)
            cRegistro.Add oCampo.CreaCampo(adInteger, , , "")
        End If
        
        'Registro.Add oCampo.CreaCampo(adInteger, , , txtPorcentajeFinanciamiento.Text)
        'Registro.Add oCampo.CreaCampo(adInteger, , , txtpagodiario.Text)
        'Registro.Add oCampo.CreaCampo(adInteger, , , txtfinan.Text)
    
    Else
    
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtnocliente.Text))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtfactura.Text))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , txtstatus.Text)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtPorcentajeFinanciamiento.Text))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtpagos.Text))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtpagodiario.Text)))
        
    End If
    cCredito.Add cRegistro
    
    Set obtenDatos = cCredito
    
'COnsideraciones:
    'Considerar todos los datos si el caption del boton cmdCredito es igual a 'Registrar'
    
    'Considerar solo los campos; Financiamiento, No. de pagos, cantidad a pagar y status si el caption del botón cmdCredito es igual a 'Actualizar'
        'Es obvio que para actualizar se debe considerar el cliente y el no. de factura.
        
End Function

Private Function imprimefn(iFactura As Long, iPagos As Integer)

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
    If iPagos = 30 Then
        oReporte.strNombreReporte = DirSys & "factura30.rpt"
    Else
        oReporte.strNombreReporte = DirSys & "factura.rpt"
    End If
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

'Private Function imprimefn(iFactura As Long)
'
'    'Declarar el uso de la clase Reporte
'    'Definir las propiedades siguientes:
'       ' EL objeto crystal report (el de la forma) sobre el cual se realizará el reporte
'       ' Definir si el reporte puede ser a pantalla o directo a impresora
'       ' Definir a que impresora se enviará el reporte
'       ' Definir el nombre del reporte (el que se diseña para la impresión en el crystal reports)
'       ' Definir los parámetros iCliente e iFactura (en una colección)
'
'    Dim oReporte As New Reporte
'    Dim cParametros As New Collection
'    Dim oCampo As New Campo
'
'    cParametros.Add oCampo.CreaCampo(adInteger, , , iFactura)  'Factura
'
'    oReporte.oCrystalReport = Me.crFactura
'    If gstrReporteEnPantalla = "Si" Then
'        oReporte.bVistaPreliminar = True
'    Else
'        oReporte.bVistaPreliminar = False
'    End If
'    oReporte.strImpresora = gPrintPed
'    oReporte.strNombreReporte = DirSys & "factura.rpt"
'    oReporte.cParametros = cParametros
'    oReporte.fnImprime
'    Set oReporte = Nothing
'
''COnsideraciones:
'
'    'Utilizar la clase Reporte
'    'antes de enviar ejecutar la impresión definir las propiedades siguientes:
'        'oCrystalReport
'        'bVistaPreliminar
'        'strImpresora
'        'strNombreReporte
'        'cParametros
'    'Ejectuar el reporte con la función fnImprime
'    'hacer el objeto reporte igual a nothing.
'
'End Function

Private Sub cmdcredito_Click()

    Dim oCredito As New credito
    'Valida los datos de la forma (la función privada de esta forma 'validaForma')
        'Si la funciòn regresa 'false' (alguno no es correcto)
            'terminar con exit sub
    If validaForma() = False Then
        Exit Sub
    End If
        
    'Si son correctos todos los datos (la funciòn 'validaForma' regreso 'true') hacer lo siguiente:
    'Si el caption del botón cmbCredito es igual a 'Registrar'
    If cmdcredito.Caption = "Registrar" Then
        'Validar la disponiblidad de crédito de un cliente dado (tomarlo del control correspondiente)
         bAltaCredito = False
        
        If True = oCredito.validaDisponibilidadDeCredito(Val(Me.txtnocliente.Text), Val(fnstrValor(txtcredito))) Then
            'Si el cliente tiene disponibilidad de crédito, mostrar la ventana 'accesofrm' para que se de la autorización.
            'accesofrm.Show vbModal
    
            'Si fue aceptada (si la propiedad bPermiteAcceso de la ventana 'accesofrm' es = true), hacer lo siguiente
            'If accesofrm.bPermiteAcceso = True Then
                'Obtener los datos de la forma
                'Registrar el nuevo crédito, enviando los datos del credito.
                
                Me.iNoFactura = oCredito.registraCredito(Val(txtnocliente.Text), _
                                                    Val(fnstrValor(txtcredito.Text)), _
                                                    Val(fnstrValor(txtpagodiario.Text)), _
                                                    Val(fnstrValor(txtfinan.Text)), _
                                                    Val(fnstrValor(txttotalpagar.Text)), _
                                                    Val(txtpagos.Text), _
                                                    dtFechaCobranzaInicial.Text, _
                                                    dtFechaCobranzaFinal.Text, _
                                                    dtFechaRegistro.Text, _
                                                    txtstatus.Text, _
                                                    lblCheque.Caption, _
                                                    lblCobrador.Caption, _
                                                    opcion.Value, _
                                                    txtdescrip.Text)
                iPagos = Val(txtpagos.Text)
            
                'Enviar mensaje indicando que ya quedó el crédito registrado
                MsgBox "Ha quedado registrado el nuevo crédito", vbInformation + vbOKOnly
                
                'If vbYes = MsgBox("¿Desea imprimir la Póliza?", vbQuestion + vbYesNo) Then
                
                    'Enviar el reporte de la factura a impresora, ejecutando la función privada 'imprimefn'
                '    Call imprimefn(iNoFactura, Val(txtPagos.Text))
                    
                'End If
                
                bAltaCredito = True
                
                Unload Me
            
            'Else
                'Si no fue aceptado (si la propiedad bPermiteAcceso de la ventana 'accesofrm' es = false)
            '    If iVeces >= 3 Then
                    'Terminar y cerrar la pantalla
            '        Unload Me
            '    End If
            'End If
        Else
            'Si el cliente NO tiene disponiblidad de crédito
            'Enviar mensaje indicando que NO tiene crédito disponible (el mensaje debe ser muy descriptivo y solo debe dar la opcion a aceptar)
            MsgBox "El cliente, ha rebasado el límite de crédito otorgado!", vbInformation + vbOKOnly
            'Dar la opción a cambiar el crédito
            txtcredito.SetFocus
        End If
    Else
    
        'En otro caso, si el caption del botón cmdCredito es igual a 'Actualizar'
        If cmdcredito.Caption = "Actualizar" Then
            
            accesofrm.Show vbModal
    
            'Si fue aceptada (si la propiedad bPermiteAcceso de la ventana 'accesofrm' es = true), hacer lo siguiente
            If accesofrm.bPermiteAcceso = True Then
                    
                Dim iFactura As Long
                
                'Obtener los datos de la forma (solo considerar: Financiamiento, No. de pagos, candtidad a pagar, status y el tipo de crédito)
                'Registrar las actualizaciones del crédito.
                iFactura = oCredito.actualizaCredito(obtenDatos)
                'Enviar el reporte de la factura a impresora, ejecutando la función privada 'fnImprime'
                Call imprimefn(Val(txtfactura.Text), Val(txtpagos.Text))
                'Enviar mensaje indicando que ya quedó el crédito actualizado
                MsgBox "Ha quedado Actualizado el crédito", vbInformation + vbOKOnly
                'Cerrar la pantalla.
            Else
                'Si no fue aceptado (si la propiedad bPermiteAcceso de la ventana 'accesofrm' es = false)
                If iVeces >= 3 Then
                    'Terminar y cerrar la pantalla
                    Unload Me
                End If
            End If
        End If
    End If
'COnsideraciones
    'Utilizar la clase 'Credito'
    'Utilizar la función validaDisponibilidadDeCredito de la clase Credito, justo para validar disponiblidad del crédito
    'Utilizar la función registraCredito de la clase Credito, para registrar el nuevo crédito. Observar que esta recibe un parámetro con todos los datos de la forma.
    'Utilizar la función actualizaCredito de la clase Credito, para realizar la actualización del crédito. Observar que esta recibe un parámetro con todos los datos de la forma.
    

End Sub

Private Sub cmdsalir_Click()
    bAltaCredito = False
    Unload Me
End Sub

Private Function calculaCredito()
    
    If bInicio = True Then
        Exit Function
    End If
    
    If Val(fnstrValor(txtcredito.Text)) <= 0# Then
        
        MsgBox "No se ha capturado la cantidad de crédito a otorgar", vbInformation, "Crédito del cliente"
        'txtcredito.SetFocus
        Exit Function
        
    End If
    
    If Val(txtPorcentajeFinanciamiento.Text) <= 0# Then
        
        MsgBox "No se ha capturado el porcentaje de financiamiento", vbInformation, "Crédito del cliente"
        txtPorcentajeFinanciamiento.SetFocus
        Exit Function
    
    End If
    
    If Val(txtpagos.Text) <= 0 Then
        
        MsgBox "No se ha capturado la cantidad de pagos", vbInformation, "Crédito del cliente"
        txtpagos.SetFocus
        Exit Function
        
    End If
    
    Dim oCredito As New credito
    Dim fPagoDiario As Double
    Dim fTotalPagar As Double
    Dim fFinanciamiento As Double
    
    'Utilizar la función 'evaluaCredito' de la clase crédito para determinar los valores siguientes:
    Call oCredito.evaluaCredito(fPagoDiario, fTotalPagar, fFinanciamiento, _
                           Val(fnstrValor(txtcredito)), Val(txtPorcentajeFinanciamiento), Val(txtpagos))
    txtpagodiario.Text = fPagoDiario
    txttotalpagar.Text = fTotalPagar
    txtfinan.Text = fFinanciamiento
                    
End Function

Private Sub cbCuentaCheques_Click()
    
    If cbCuentaCheques.ListIndex <> -1 Then
        
        Dim oCuentaCheques As New CuentaCheques
        Dim dSaldoCuenta As Double
        Dim iChequeDisponible As Long
        Dim iCuenta As Integer
        
        iCuenta = cbCuentaCheques.ItemData(cbCuentaCheques.ListIndex)
        
        Call oCuentaCheques.saldoCuenta(iCuenta, dSaldoCuenta)
        
        txtSaldoCuenta = dSaldoCuenta
               
        If oCuentaCheques.siguienteDisponible(iCuenta, iChequeDisponible) = True Then
            lblCheque.Caption = iChequeDisponible
        Else
            If vbYes = MsgBox("De la cuenta seleccionada, no hay cheques disponibles." & Chr(13) & "Para registrar cheques seleccione Yes" & Chr(13) & "Si desea por ahora hacer el pago por transferencia, seleccione No", vbInformation + vbYesNo) Then
                
                mantenimientoChequesfrm.iCuentaCheques = iCuenta
                mantenimientoChequesfrm.strCuentaCheques = cbCuentaCheques.Text
                mantenimientoChequesfrm.Show vbModal
                If oCuentaCheques.siguienteDisponible(iCuenta, iChequeDisponible) = True Then
                    lblCheque.Caption = iChequeDisponible
                End If
            End If
        End If
        
        Set oCuentaCheques = Nothing
        
    End If
    
End Sub

Private Sub opcion_Click()
    
    If opcion.Value = 1 Then
        
        txtdescrip.Text = ""
        txtdescrip.Visible = True
        lblDescripcion.Visible = True
        txtPorcentajeFinanciamiento.Text = 0
        txtpagos.Text = 60
    
    ElseIf opcion.Value = 0 Then
        
        txtdescrip.Text = ""
        txtdescrip.Visible = False
        lblDescripcion.Visible = False
        txtPorcentajeFinanciamiento.Text = 14
        txtpagos.Text = 30
        
    End If
    
    txtcredito.SetFocus
    
End Sub

Private Sub txtCredito_Change()
    'Ejecuta la función 'calculaCredito' (privada de esta forma) para que se actualice el calculo del crédito
    calculaCredito
End Sub

'Private Sub txtpagos_LostFocus()
'
'    'Ejecuta la función 'calculaCredito' (privada de esta forma) para que se actualice el calculo del crédito
'    If Val(txtPagos) <> 30 And Val(txtPagos) <> 60 Then
'
'        MsgBox "Los plazos válidos son 30 o 60 días, por favor corrija"
'
'        txtPagos.SetFocus
'
'    End If
'
'    calculaCredito
'
'    dtFechaCobranzaFinal = DateAdd("d", txtPagos.Text, dtFechaCobranzaInicial)
'
'End Sub

Private Sub txtPorcentajeFinanciamiento_Change()
    'Ejecuta la función 'calculaCredito' (privada de esta forma) para que se actualice el calculo del crédito
    calculaCredito
End Sub

Private Sub txtPagos_Change()
    'Ejecuta la función 'calculaCredito' (privada de esta forma) para que se actualice el calculo del crédito
    calculaCredito
    If Len(txtpagos.Text) > 0 Then
        dtFechaCobranzaFinal = DateAdd("d", txtpagos.Text, dtFechaCobranzaInicial)
    End If
End Sub

