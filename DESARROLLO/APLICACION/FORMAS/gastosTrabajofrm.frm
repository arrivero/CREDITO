VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form gastosTrabajofrm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Gastos de Trabajo"
   ClientHeight    =   6960
   ClientLeft      =   3060
   ClientTop       =   690
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlMovimientos 
      Height          =   6645
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   420
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   11721
      _Version        =   196608
      BackColor       =   12648447
      BevelWidth      =   0
      BorderWidth     =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gastos Registrados"
         Height          =   4935
         Left            =   45
         TabIndex        =   8
         Top             =   1560
         Width           =   5535
         Begin VB.CommandButton cmdgraba 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Registra Gastos"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2850
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   4200
            Width           =   1335
         End
         Begin VB.CommandButton cmdMuestra 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Mostrar Gastos"
            Height          =   495
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   4200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin FPSpread.vaSpread sprPago 
            Height          =   3195
            Left            =   60
            TabIndex        =   9
            Top             =   300
            Width           =   5415
            _Version        =   196608
            _ExtentX        =   9551
            _ExtentY        =   5636
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   13106686
            GridColor       =   8454143
            MaxCols         =   4
            MaxRows         =   20
            OperationMode   =   2
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            ShadowColor     =   13106686
            SpreadDesigner  =   "gastosTrabajofrm.frx":0000
         End
         Begin Crystal.CrystalReport crGastosTrabajo 
            Left            =   720
            Top             =   4260
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label txtTotalGastos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   3330
            TabIndex        =   29
            Top             =   3690
            Width           =   2040
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Total de Gastos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   630
            TabIndex        =   10
            Top             =   3660
            Width           =   2565
         End
      End
      Begin VB.Frame fraGasto 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gastos"
         Height          =   1575
         Left            =   30
         TabIndex        =   11
         Top             =   0
         Width           =   5535
         Begin MSMask.MaskEdBox txtimporte 
            Height          =   330
            Left            =   1080
            TabIndex        =   28
            Top             =   585
            Width           =   1815
            _ExtentX        =   3201
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
            Mask            =   "$#####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtgasto 
            BackColor       =   &H00C0FFFF&
            Height          =   330
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   26
            Top             =   180
            Width           =   2805
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox combo 
            Height          =   315
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdagrega 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Agrega Gasto"
            Height          =   495
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbgasto 
            Height          =   315
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin EditLib.fpCurrency txtimporteold 
            Height          =   315
            Left            =   1080
            TabIndex        =   2
            Top             =   600
            Visible         =   0   'False
            Width           =   1845
            _Version        =   196608
            _ExtentX        =   3254
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
            BackColor       =   12648447
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
            CurrencyDecimalPlaces=   2
            CurrencyNegFormat=   0
            CurrencyPlacement=   0
            CurrencySymbol  =   "$"
            DecimalPoint    =   "."
            FixedPoint      =   -1  'True
            LeadZero        =   0
            MaxValue        =   "20000"
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
         Begin SSCalendarWidgets_A.SSDateCombo txtfechagasto 
            Height          =   315
            Left            =   1080
            TabIndex        =   3
            Top             =   960
            Width           =   1845
            _Version        =   65537
            _ExtentX        =   3254
            _ExtentY        =   556
            _StockProps     =   93
            Format          =   "DD/MM/YY"
            BevelColorFace  =   12648447
            Mask            =   2
         End
         Begin Crystal.CrystalReport reppagos 
            Left            =   3660
            Top             =   990
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            ReportFileName  =   "C:\Facturas\repgastos.rpt"
            PrintFileLinesPerPage=   60
         End
         Begin EditLib.fpText txtgastoold 
            Height          =   315
            Left            =   1080
            TabIndex        =   1
            Top             =   210
            Visible         =   0   'False
            Width           =   2805
            _Version        =   196608
            _ExtentX        =   4948
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
            BackColor       =   12648447
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
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   0
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
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin VB.Label lblModificaGastos 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Seleccione Fecha y enter para modificar gastos."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   1290
            Visible         =   0   'False
            Width           =   3585
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   1005
            Width           =   855
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Importe:"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Gasto:"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   285
            Width           =   735
         End
      End
   End
   Begin Threed.SSPanel pnlMovimientos 
      Height          =   6645
      Index           =   1
      Left            =   0
      TabIndex        =   19
      Top             =   420
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11721
      _Version        =   196608
      BackColor       =   12648447
      BevelWidth      =   0
      BorderWidth     =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
      Begin MSMask.MaskEdBox txtImporteDeposito 
         Height          =   330
         Left            =   1395
         TabIndex        =   27
         Top             =   990
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "$#####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdguarda 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Registrar"
         Height          =   495
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1470
         Width           =   1215
      End
      Begin EditLib.fpCurrency txtImporteDepositoold 
         Height          =   345
         Left            =   1380
         TabIndex        =   24
         Top             =   990
         Width           =   1845
         _Version        =   196608
         _ExtentX        =   3254
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
         BackColor       =   12648447
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
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   -1
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   ""
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "100000"
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
      Begin SSCalendarWidgets_A.SSDateCombo txtfecha 
         Height          =   315
         Left            =   1380
         TabIndex        =   23
         Top             =   540
         Width           =   1845
         _Version        =   65537
         _ExtentX        =   3254
         _ExtentY        =   556
         _StockProps     =   93
         BevelColorFace  =   12648447
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Importe:"
         Height          =   225
         Left            =   420
         TabIndex        =   22
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha:"
         Height          =   225
         Left            =   450
         TabIndex        =   21
         Top             =   570
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salir"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6180
      Width           =   1215
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   360
      Top             =   7350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196608
      MinFontSize     =   1
      MaxFontSize     =   100
   End
   Begin ComctlLib.TabStrip tabMovimientos 
      Height          =   6945
      Left            =   -15
      TabIndex        =   0
      Top             =   45
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   12250
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Gastos"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Depósitos"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   7290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "gastosTrabajofrm.frx":0494
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "gastosTrabajofrm.frx":066E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "gastosTrabajofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_DESCRIPCION = 1
Private Const COL_IMPORTE = 2
Private Const COL_FECHA = 3
Private Const COL_USUARIO = 4

Private miFrameActivo As Integer

Private bModifica As Boolean

Private iTrabajoInternos As Integer '1 - Gastos de Trabajo   0 - Gastos Internos

Dim renglongto As Long

Private Sub cmdguarda_Click()
    
    Dim ImporteDeposito As Double
    ImporteDeposito = Val(fnstrValor(txtImporteDeposito.ClipText))
    
    If ImporteDeposito > 0# Then
        
        Dim oGasto As New Gasto
        oGasto.deposito txtfecha.Text, ImporteDeposito
        Set oGasto = Nothing
        
        MsgBox "El depósito ha sido registrado", vbInformation + vbOKOnly, "Depósitos"
        txtImporteDeposito.SetFocus
        
        'txtImporteDeposito.Mask = ""
        txtImporteDeposito.Text = ""
        'txtImporteDeposito.Mask = "$##,###.##"
        
    Else
        MsgBox "El importe no es correcto, verfique por favor!", vbInformation + vbOKOnly
        txtImporteDeposito.SetFocus
    End If

End Sub

Private Sub Form_Activate()

    If giTipoUsuario = USUARIO_GERENTE Then
        lblModificaGastos.Visible = True
    Else
        tabMovimientos.Enabled = False
    End If
    
    If iTrabajoInternos = 1 Then
        fraGasto.Caption = "Gastos " + UCase(gstrUsuario)
        cmdMuestra.Visible = True
    Else
        fraGasto.Caption = "Gastos Internos"
    End If

End Sub

Private Sub Form_Load()
        
    gastosTrabajoInternosfrm.Show vbModal
    iTrabajoInternos = gastosTrabajoInternosfrm.iTrabajoInternos
        
    txtfecha.Text = Format(Now, "dd/mm/yyyy")
        
    txtfechagasto.Text = Format(Now, "dd/mm/yyyy")
    
    txtfechagasto_KeyPress (13)
    
End Sub

Private Function bajaCaptura(lRow As Long)

    Dim fPago As Double
    'fPago = Val(fnstrValor(txtimporte.Text))
    fPago = Val(fnstrValor(txtimporte.ClipText))
        
    If fPago > 0# Then
        
        If cmdagrega.Caption = "Agrega Gasto" Then
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
    
        sprPago.Col = COL_DESCRIPCION
        sprPago.Text = txtgasto.Text
        sprPago.Col = COL_IMPORTE
        'sprPago.Text = fnstrValor(txtimporte.Text)
        sprPago.Text = fnstrValor(txtimporte)
        sprPago.Col = COL_FECHA
        sprPago.Text = txtfechagasto.Text
        
        If iTrabajoInternos = 1 Then
            sprPago.Col = COL_USUARIO
            sprPago.Text = gstrUsuario
        End If
        
        txtTotalGastos.Caption = Format(obtenTotalGrid(sprPago, COL_IMPORTE), "$ ###,###,###,###0.00")
        
        txtgasto.Text = ""
        txtimporte.Text = ""
        'txtimporte.Mask = ""
        'txtimporte.Text = ""
        'txtimporte.Mask = "$##,###.##"
        
        cmdgraba.Enabled = True
        
        txtgasto.SetFocus

    Else
        MsgBox "¡El monto del gasto debe ser mayor a $0.0 pesos, verifique por favor!", vbInformation + vbOKOnly, "Registro de Pagos"
        txtgasto.SetFocus
    End If
        
End Function

Private Sub sprPago_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If giTipoUsuario = USUARIO_GERENTE Then
        sprPago.ToolTipText = "¡ Del o F7 para eliminar el gasto !"
    End If
    
End Sub

Private Sub sprPago_KeyUp(KeyCode As Integer, Shift As Integer)

      Select Case KeyCode
          
          Case vbKeyDelete, vbKeyF7
            
            If giTipoUsuario = USUARIO_GERENTE Then
            
                If vbOK = MsgBox("Esta seguro de eliminar el gasto?", vbQuestion + vbOKCancel) Then
                  sprPago.Row = sprPago.ActiveRow
                  sprPago.Action = ActionDeleteRow
                  txtTotalGastos.Caption = Format(obtenTotalGrid(sprPago, COL_IMPORTE), "###,###,###,###0.00")
                  cmdgraba.Enabled = True
                  bModifica = True
                
                End If
            End If
            
      End Select

End Sub

Private Sub sprPago_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If giTipoUsuario = USUARIO_GERENTE Then
    
        sprPago.Row = Row
        sprPago.Col = COL_DESCRIPCION
        
        renglongto = Row
        
        If sprPago.Text <> "" Then
            txtgasto.Text = sprPago.Text
            sprPago.Col = COL_IMPORTE
            txtimporte.Text = fnstrValor(sprPago.Text)
            sprPago.Col = COL_FECHA
            txtfechagasto.Text = sprPago.Text
        End If
        
        cmdagrega.Caption = "Modifica Gasto"
        
    End If
    
End Sub

Private Sub cmdagrega_Click()

    If cmdagrega.Caption = "Modifica Gasto" Then
        bajaCaptura renglongto
        cmdagrega.Caption = "Agrega Gasto"
    ElseIf cmdagrega.Caption <> "Modifica Gasto" Then
        bajaCaptura 1
    End If
        
End Sub

Private Sub cmdgraba_Click()
    
    Dim oGasto As New Gasto
    If bModifica = False Then
        If iTrabajoInternos = 1 Then
            oGasto.registraGasto obtenGastos(0)
        Else
            oGasto.registraGasto obtenGastos(2)
        End If
        MsgBox "Los cambios han sido registrados", vbOKOnly + vbInformation, "Gastos"
    
    Else
        If iTrabajoInternos = 1 Then
            oGasto.actualizaGasto txtfechagasto.Text, obtenGastos(0)
        Else
            oGasto.actualizaGastoInterno txtfechagasto.Text, obtenGastos(2)
        End If
        MsgBox "!Los gastos fueron registrados¡", vbOKOnly + vbInformation
    
    End If
    Set oGasto = Nothing

End Sub

Private Function obtenGastos(iReporteRegistra As Integer) As Collection

    Dim cGastos As New Collection
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    
    Dim lRow As Long
    Dim iGasto As Integer
    
    For lRow = 1 To sprPago.DataRowCnt
    
        Set cRegistro = New Collection
        
        sprPago.Row = lRow
           
        sprPago.Col = COL_DESCRIPCION
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text) 'Descripción del gasto
        sprPago.Col = COL_IMPORTE
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPago.Text))) 'iMPORTE DEL GASTO
        sprPago.Col = COL_FECHA
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text) 'fECHA
        
        If iTrabajoInternos = 1 Then
            cRegistro.Add oCampo.CreaCampo(adInteger, , , iGasto) '#Pago
            sprPago.Col = COL_USUARIO
            cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago.Text) 'usuario
        Else
            cRegistro.Add oCampo.CreaCampo(adInteger, , , 0) '#Pago
            cRegistro.Add oCampo.CreaCampo(adInteger, , , "") 'usuario
        End If
        
        cRegistro.Add oCampo.CreaCampo(adInteger, , , iReporteRegistra) 'Reporte = 1, Registra = 0, Interno = 2
        
        cGastos.Add cRegistro
        iGasto = iGasto + 1
        
    Next lRow
    
    Set obtenGastos = cGastos
    
End Function

Private Sub fnImprime(strReporte As String, crObjeto As CrystalReport)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    Dim strNombreReporte As String
    
    strNombreReporte = strReporte + ".rpt"
    
    cParametros.Add oCampo.CreaCampo(adInteger, , , Format(Now, "dd/mm/yyyy"))

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

Private Sub cmdMuestra_Click()

    Dim oGasto As New Gasto
    oGasto.registraGasto obtenGastos(1)
    Set oGasto = Nothing
    fnImprime "gastosTrabajo", crGastosTrabajo
    
End Sub

Private Sub tabMovimientos_Click()
    
    If giTipoUsuario = USUARIO_GERENTE Then

        If tabMovimientos.SelectedItem.Index - 1 = miFrameActivo Then Exit Sub ' No need to change frame.
        
        ' Comosea, oculta el frame anterior, muestra el nuevo.
        pnlMovimientos(tabMovimientos.SelectedItem.Index - 1).Visible = True
        pnlMovimientos(miFrameActivo).Visible = False
        
        miFrameActivo = tabMovimientos.SelectedItem.Index - 1
        
    End If

End Sub

Private Sub txtfechagasto_KeyPress(KeyAscii As Integer)
    
        If KeyAscii = 13 Then
        
            'Limpia grid
            fnLimpiaGrid sprPago
            
            Dim oGasto As New Gasto
            oGasto.obtenGastos txtfechagasto.Text, iTrabajoInternos
            
            'llena el grid con los gastos
            fnLlenaTablaCollection sprPago, oGasto.cDatos
            
            'Calcula el total de los gastos
            txtTotalGastos.Caption = Format(obtenTotalGrid(sprPago, COL_IMPORTE), "###,###,###,###0.00")
            
            Set oGasto = Nothing
        
            bModifica = True
            
        End If
    
End Sub

Private Sub cmdsalir_Click()
    sicPrincipalfrm.Caption = ""
    sicPrincipalfrm.pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
    despliegaVentana portadafrm, WND_PORTADA
End Sub

