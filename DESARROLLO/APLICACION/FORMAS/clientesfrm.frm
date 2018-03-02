VERSION 5.00
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form clientesfrm 
   BorderStyle     =   0  'None
   ClientHeight    =   10725
   ClientLeft      =   315
   ClientTop       =   30
   ClientWidth     =   13020
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10725
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196608
      MinFontSize     =   1
      MaxFontSize     =   100
   End
   Begin VB.Frame Frame6 
      Caption         =   "Créditos"
      ForeColor       =   &H8000000D&
      Height          =   3675
      Left            =   30
      TabIndex        =   35
      Top             =   7020
      Width           =   12945
      Begin Crystal.CrystalReport crFactura 
         Left            =   12420
         Top             =   1560
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdImprimePoliza 
         Caption         =   "Póliza"
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
         Left            =   11520
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCargaCreditos 
         Caption         =   "Carga Automática"
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
         Height          =   525
         Left            =   11520
         TabIndex        =   17
         ToolTipText     =   "Del credito seleccionado, modifica pagos."
         Top             =   240
         Width           =   1335
      End
      Begin Crystal.CrystalReport crSaldos 
         Left            =   12420
         Top             =   3090
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdSaldosIndividuales 
         Caption         =   "Saldos Individuales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
         TabIndex        =   22
         Top             =   2970
         Width           =   1335
      End
      Begin VB.CommandButton cmdModificaCredito 
         Caption         =   "Modifica Crédito"
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
         Left            =   11520
         TabIndex        =   20
         Top             =   2010
         Width           =   1335
      End
      Begin Threed.SSPanel pnlCredito 
         Height          =   2865
         Left            =   60
         TabIndex        =   46
         Top             =   720
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   5054
         _Version        =   196608
         Caption         =   "pnlCredito(0)"
         RoundedCorners  =   0   'False
         FloodShowPct    =   0   'False
         Begin FPSpread.vaSpread sprCredito 
            Height          =   2835
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   11355
            _Version        =   196608
            _ExtentX        =   20029
            _ExtentY        =   5001
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
            SpreadDesigner  =   "clientesfrm.frx":0000
         End
      End
      Begin ComctlLib.TabStrip tabCredito 
         Height          =   3345
         Left            =   30
         TabIndex        =   45
         Top             =   270
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5900
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   5
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Vigentes"
               Key             =   ""
               Object.Tag             =   """V"""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Pendientes"
               Key             =   ""
               Object.Tag             =   """P"""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Cancelados"
               Key             =   ""
               Object.Tag             =   """C"""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Vencidos"
               Key             =   ""
               Object.Tag             =   """V"""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Terminados"
               Key             =   ""
               Object.Tag             =   """T"""
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
      Begin VB.CommandButton cmdpagos 
         Caption         =   "Modifica Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   11520
         TabIndex        =   21
         Top             =   2490
         Width           =   1335
      End
      Begin VB.CommandButton cmdcredito 
         Caption         =   "Nuevo crédito"
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
         Height          =   435
         Left            =   11520
         TabIndex        =   18
         Top             =   780
         Width           =   1335
      End
   End
   Begin VB.Frame fraClientes 
      Caption         =   "Lista de clientes"
      ForeColor       =   &H00FF0000&
      Height          =   7005
      Left            =   5430
      TabIndex        =   34
      Top             =   30
      Width           =   7545
      Begin Threed.SSPanel pnlListaClientes 
         Height          =   6615
         Left            =   30
         TabIndex        =   49
         Top             =   330
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   11668
         _Version        =   196608
         RoundedCorners  =   0   'False
         FloodShowPct    =   0   'False
         Begin FPSpread.vaSpread sprCliente 
            Height          =   6615
            Left            =   -1680
            TabIndex        =   52
            Top             =   30
            Width           =   7755
            _Version        =   196608
            _ExtentX        =   13679
            _ExtentY        =   11668
            _StockProps     =   64
            AutoCalc        =   0   'False
            AutoClipboard   =   0   'False
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
            FormulaSync     =   0   'False
            GrayAreaBackColor=   16777215
            MaxCols         =   3
            OperationMode   =   2
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "clientesfrm.frx":1C12
            Appearance      =   2
         End
      End
      Begin Threed.SSPanel pnlCreditosNuevos 
         Height          =   6615
         Left            =   30
         TabIndex        =   50
         Top             =   330
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   11668
         _Version        =   196608
         RoundedCorners  =   0   'False
         FloodShowPct    =   0   'False
         Begin FPSpread.vaSpread sprCreditoNuevo 
            Height          =   6555
            Left            =   30
            TabIndex        =   51
            Top             =   30
            Width           =   6015
            _Version        =   196608
            _ExtentX        =   10610
            _ExtentY        =   11562
            _StockProps     =   64
            AutoClipboard   =   0   'False
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
            MaxCols         =   6
            OperationMode   =   1
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "clientesfrm.frx":358F
            Appearance      =   2
         End
      End
      Begin Crystal.CrystalReport crClientesVencidos 
         Left            =   6990
         Top             =   5940
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Threed.SSFrame fraReportes 
         Height          =   3555
         Left            =   6120
         TabIndex        =   47
         Top             =   3420
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   6271
         _Version        =   196608
         Caption         =   "Reportes"
         Alignment       =   2
         Begin VB.CommandButton cmdPagosUsuario 
            Caption         =   "Pagos por Usuario"
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
            Height          =   435
            Left            =   60
            TabIndex        =   29
            Top             =   3060
            Width           =   1245
         End
         Begin Crystal.CrystalReport crClientesVigentes 
            Left            =   870
            Top             =   420
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin Crystal.CrystalReport crClientesTotal 
            Left            =   870
            Top             =   2100
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin Crystal.CrystalReport crClientesCancelados 
            Left            =   870
            Top             =   1650
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin Crystal.CrystalReport crClientesTerminados 
            Left            =   870
            Top             =   1230
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin Crystal.CrystalReport crClientesPendientes 
            Left            =   870
            Top             =   840
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.CommandButton cmdTotal 
            Caption         =   "Total"
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
            Height          =   375
            Left            =   60
            TabIndex        =   28
            Top             =   2490
            Width           =   1245
         End
         Begin VB.CommandButton cmdVencidos 
            Caption         =   "Vencidos"
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
            Height          =   375
            Left            =   60
            TabIndex        =   27
            Top             =   2040
            Width           =   1245
         End
         Begin VB.CommandButton cmdCancelados 
            Caption         =   "Cancelados"
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
            Height          =   375
            Left            =   60
            TabIndex        =   26
            Top             =   1650
            Width           =   1245
         End
         Begin VB.CommandButton cmdTerminados 
            Caption         =   "Terminados"
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
            Height          =   375
            Left            =   60
            TabIndex        =   25
            Top             =   1260
            Width           =   1245
         End
         Begin VB.CommandButton cmdPendientes 
            Caption         =   "Pendientes"
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
            Height          =   375
            Left            =   60
            TabIndex        =   24
            Top             =   870
            Width           =   1245
         End
         Begin VB.CommandButton cmdVigentes 
            Caption         =   "Vigentes"
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
            Height          =   375
            Left            =   60
            TabIndex        =   23
            Top             =   480
            Width           =   1245
         End
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
         Height          =   375
         Left            =   6150
         TabIndex        =   30
         Top             =   1260
         Width           =   1335
      End
      Begin VB.CommandButton cmdgrabar 
         Caption         =   "Alta"
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
         Height          =   375
         Left            =   6150
         TabIndex        =   16
         Top             =   630
         Width           =   1335
      End
      Begin VB.CommandButton cmdnuevo 
         Caption         =   "Nuevo"
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
         Height          =   375
         Left            =   6150
         TabIndex        =   15
         Top             =   210
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog cmdListaCreditos 
         Left            =   7020
         Top             =   1710
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontSize        =   10
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      ForeColor       =   &H00FF0000&
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5385
      Begin VB.TextBox txtApellidoMaterno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1485
         Width           =   4095
      End
      Begin MSMask.MaskEdBox txtnocliente 
         Height          =   330
         Left            =   1215
         TabIndex        =   1
         Top             =   360
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame4 
         Caption         =   "Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   90
         TabIndex        =   42
         Top             =   5865
         Width           =   5175
         Begin VB.ComboBox cbCobrador 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   675
            Width           =   2415
         End
         Begin MSMask.MaskEdBox txtmaxcredito 
            Height          =   375
            Left            =   1395
            TabIndex        =   12
            Top             =   180
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            PromptInclude   =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "$ ######"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtAtrasos 
            Height          =   375
            Left            =   4005
            TabIndex        =   13
            Top             =   225
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            Mask            =   "#"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
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
            Left            =   225
            TabIndex        =   56
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Monto máximo de crédito:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   240
            TabIndex        =   44
            Top             =   195
            Width           =   1050
         End
         Begin VB.Label lbatrasos 
            Caption         =   "Atrasos Permitidos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2880
            TabIndex        =   43
            Top             =   195
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dirección y Teléfono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   90
         TabIndex        =   36
         Top             =   1965
         Width           =   5205
         Begin MSMask.MaskEdBox txttelefono 
            Height          =   375
            Left            =   2835
            TabIndex        =   10
            Top             =   2070
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##-#### ####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtcp 
            Height          =   375
            Left            =   630
            TabIndex        =   9
            Top             =   2070
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   661
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtEntidadFederativa 
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
            Left            =   1335
            MaxLength       =   100
            TabIndex        =   8
            Top             =   1620
            Width           =   2415
         End
         Begin VB.TextBox txtColonia 
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
            Left            =   1335
            MaxLength       =   100
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   720
            Width           =   3765
         End
         Begin VB.TextBox txtcd 
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
            Left            =   1335
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1170
            Width           =   2415
         End
         Begin VB.TextBox txtubicacion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   150
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   2790
            Width           =   4980
         End
         Begin VB.TextBox txtdireccion 
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
            Left            =   1335
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   270
            Width           =   3765
         End
         Begin VB.Label Label14 
            Caption         =   "Estado:"
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
            TabIndex        =   55
            Top             =   1665
            Width           =   825
         End
         Begin VB.Label Label12 
            Caption         =   "Colonia:"
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
            TabIndex        =   54
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label9 
            Caption         =   "Teléfono:"
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
            Left            =   2025
            TabIndex        =   41
            Top             =   2115
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Municipio:"
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
            TabIndex        =   40
            Top             =   1215
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "C.P.:"
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
            TabIndex        =   39
            Top             =   2115
            Width           =   600
         End
         Begin VB.Label Label5 
            Caption         =   "Descripción de Ubicación:"
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
            TabIndex        =   38
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Calle y Número:"
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
            TabIndex        =   37
            Top             =   315
            Width           =   1185
         End
      End
      Begin VB.TextBox txtnombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1215
         MaxLength       =   100
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtapellido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1215
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label11 
         Caption         =   "Ap Materno:"
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
         Left            =   225
         TabIndex        =   53
         Top             =   1515
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "No. Cliente:"
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
         Left            =   240
         TabIndex        =   48
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre(s):"
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
         Left            =   240
         TabIndex        =   33
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Ap Paterno:"
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
         Left            =   240
         TabIndex        =   32
         Top             =   1110
         Width           =   915
      End
   End
End
Attribute VB_Name = "clientesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim paso As Integer

Private iTabActivo As Integer
Private iRowActual_creditos As Integer

Private Const CREDITO_FACTURA = 1
Private Const TAB_VIGENTES = 1
Private Const TAB_PENDIENTES = 2
Private Const TAB_CANCELADOS = 3
Private Const TAB_VENCIDOS = 4
Private Const TAB_TERMINADOS = 5

Private Const COL_CLIENTE_ID = 1

'Private Const COL_CREDITOS_SELECCIONADO = 1
Private Const COL_CREDITOS_CLIENTE = 1
Private Const COL_CREDITOS_NOMBRE_CLIENTE = 2
Private Const COL_CREDITOS_CREDITO = 3
'Private Const COL_CREDITOS_INTERES = 5
'Private Const COL_CREDITOS_FINANCIAMIENTO = 6
'Private Const COL_CREDITOS_NO_PAGOS = 7
'Private Const COL_CREDITOS_CANT_PAGAR = 8
'Private Const COL_CREDITOS_TOTAL_PAGAR = 9
Private Const COL_CREDITOS_AUTORIZADO = 4

Private iVeces As Integer

Private bCreditosNuevos As Boolean

Dim cCreditos As New Collection

Private Sub cmdImprimePoliza_Click()

    sprCredito.Col = 1 'COL_FOLIO
    sprCredito.Row = sprCredito.ActiveRow
    
    If Val(sprCredito.Text) <= 0 Then
        MsgBox "Seleccione un cliente y su crédito", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    Dim oCredito As New credito
        
    If oCredito.datosCredito(Val(sprCredito.Text)) = True Then
    
        Dim cRegistros As New Collection
        Dim cRegistro As New Collection
        Dim oCampo As New Campo
        
        Set cRegistros = oCredito.cDatos
        Set cRegistro = cRegistros(1)
        Set oCampo = cRegistro(7)
        
        Call imprimefn(Val(sprCredito.Text), oCampo.Valor)
        
    End If
    
    Set oCredito = Nothing
    
End Sub

Private Sub cmdPagosUsuario_Click()
    pagosPorCobrador.Show vbModal
End Sub

Private Sub Form_Activate()
    txtnocliente.SetFocus
End Sub
    
Private Sub Form_Load()


    Select Case giTipoUsuario
        Case Is = USUARIO_GERENTE
            cmdNuevo.Enabled = True
            cmdgrabar.Enabled = True
            
            cmdModificaCredito.Enabled = True
            
            cmdVigentes.Enabled = True
            cmdPendientes.Enabled = True
            cmdTerminados.Enabled = True
            cmdCancelados.Enabled = True
            cmdVencidos.Enabled = True
            cmdTotal.Enabled = True
            cmdCargaCreditos.Enabled = True
            cmdPagosUsuario.Enabled = True
        
        Case Is = USUARIO_ADMINSTRADOR
            cmdVigentes.Enabled = True
            cmdPendientes.Enabled = True
            cmdTerminados.Enabled = True
            cmdCancelados.Enabled = True
            cmdVencidos.Enabled = True
            cmdTotal.Enabled = True
            cmdCargaCreditos.Enabled = True
            cmdPagosUsuario.Enabled = True
        'Case Is = USUARIO_USUARIO
    
    End Select
    
    Dim strUsuario As String
    strUsuario = gstrUsuario

    Dim oCliente As New Cliente
    
    If oCliente.listaLimiteCredito Then
        
        Call fnLlenaTablaCollection(sprCliente, oCliente.cDatos)
        
    End If
    Set oCliente = Nothing
    
    iTabActivo = TAB_VIGENTES
    
    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        fnLlenaComboCollecion cbCobrador, oUsuario.cDatos, 0, ""
    End If
    Set oUsuario = Nothing
        
    bCreditosNuevos = False
    
End Sub

Private Function validaForma(bActualiza As Boolean) As Boolean
    
    Dim strMsgEncabezado As String
    
    validaForma = False
    
    
    If bActualiza = True Then
    
        strMsgEncabezado = "Actualiza datos de Cliente"
        
        If Me.txtnocliente.Text = "" Then
            MsgBox "No se ha capturado el No. del cliente", vbOKOnly, strMsgEncabezado
            txtNombre.SetFocus
        End If
    Else
        strMsgEncabezado = "Alta de Cliente"
    End If
    
    If txtNombre.Text = "" Then
        MsgBox "No se ha capturado el nombre del cliente", vbOKOnly, strMsgEncabezado
        txtNombre.SetFocus
    End If
        
    If txtapellido.Text = "" Then
        MsgBox "No se ha capturado el apellido paterno del cliente", vbOKOnly, strMsgEncabezado
        txtapellido.SetFocus
        Exit Function
    End If
    
    If txtApellidoMaterno.Text = "" Then
        MsgBox "No se ha capturado el apellido materno", vbOKOnly, strMsgEncabezado
        txtApellidoMaterno.SetFocus
        Exit Function
    End If
    
    If txtdireccion.Text = "" Then
        MsgBox "No se ha capturado la calle y el número", vbOKOnly, strMsgEncabezado
        txtdireccion.SetFocus
        Exit Function
    End If
    
    If txtColonia.Text = "" Then
        MsgBox "No se ha capturado la colonia", vbOKOnly, strMsgEncabezado
        txtColonia.SetFocus
        Exit Function
    End If
    
    If txtcd.Text = "" Then
        MsgBox "No se ha capturado el municipio", vbOKOnly, strMsgEncabezado
        txtcd.SetFocus
        Exit Function
    End If
    
    If txtEntidadFederativa.Text = "" Then
        MsgBox "No se ha capturado el Estado", vbOKOnly, strMsgEncabezado
        txtEntidadFederativa.SetFocus
        Exit Function
    End If
    
    If cbCobrador.Text = "" Then
        MsgBox "No se ha capturado el Cobrador", vbOKOnly, strMsgEncabezado
        cbCobrador.SetFocus
        Exit Function
    End If
    
    If (Val(txtmaxcredito.Text) / 100#) < 0# Then
        MsgBox "No se ha capturado el límite de crédito del cliente", vbOKOnly, strMsgEncabezado
        txtmaxcredito.SetFocus
        Exit Function
    End If
    
    If txtAtrasos.Text = "" Then
        MsgBox "No se han capturado los atrasos permitidos del cliente", vbOKOnly, strMsgEncabezado
        Exit Function
    End If
    
    validaForma = True
    
End Function

Private Function obtenDatos(bActualiza As Boolean) As Collection

    'no_cliente , Nombre, apellido, DIRECCION, ubicacion, cp, ciudad, estado, TELEFONO, maxcredito, atrasospermitidos
    
    Dim cCliente As New Collection
    Dim Registro As New Collection
    Dim oCampo As New Campo

    If bActualiza = True Then
        Registro.Add oCampo.CreaCampo(adInteger, , , Val(Me.txtnocliente.Text))
    End If
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtNombre)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtapellido)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtdireccion)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtubicacion)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtcp)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtcd)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.cbCobrador.Text)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtTelefono)
    Registro.Add oCampo.CreaCampo(adInteger, , , Val(txtmaxcredito.Text))
    Registro.Add oCampo.CreaCampo(adInteger, , , Val(txtAtrasos.Text))
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtApellidoMaterno.Text)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtColonia.Text)
    Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtEntidadFederativa.Text)
        
    cCliente.Add Registro

    Set obtenDatos = cCliente
    
End Function

Private Sub cmdCargaCreditos_Click()
    
    If bCreditosNuevos = True Then
        fraClientes.Caption = "Lista de Clientes"
        pnlListaClientes.Visible = True
        pnlCreditosNuevos.Visible = False
        bCreditosNuevos = False
        cmdCargaCreditos.Caption = "Carga Automática"
    Else
        
'        If gEncriptado = "NO" Then
            If cargaCreditosHH = True Then
            
                fraClientes.Caption = "Lista de Creditos Nuevos"
                pnlListaClientes.Visible = False
                pnlCreditosNuevos.Visible = True
                bCreditosNuevos = True
                cmdCargaCreditos.Caption = "Autoriza Créditos"
                
            End If
'        Else
'            Dim lResultado As Long
'
'            lResultado = ShellExecute(Me.hwnd, "", App.Path & "\Decrypt\Decrypt.exe", "", "", 3)
'
'            If lResultado > 32 Then
'
'                MsgBox "Ahora se van a cargar los creditos solicitados.", vbOKOnly
'                'MsgBox "Antes de continuar, primero seleccione el achivo a integrar", vbOKOnly
'
'                If cargaCreditosHH = True Then
'
'                    fraClientes.Caption = "Lista de Creditos Nuevos"
'                    pnlListaClientes.Visible = False
'                    pnlCreditosNuevos.Visible = True
'                    bCreditosNuevos = True
'                    cmdCargaCreditos.Caption = "Autoriza Créditos"
'
'                End If
'
'            End If
            
'        End If
        
    End If
    
    cmdCargaCreditos.SetFocus
    
End Sub

Private Function imprimefn(iFactura As Long, iPagos As Integer)

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
    If iPagos <= 30 Then
        oReporte.strNombreReporte = DirSys & "factura30.rpt"
    Else
        oReporte.strNombreReporte = DirSys & "factura.rpt"
    End If
    oReporte.cParametros = cParametros
    oReporte.fnImprime
    Set oReporte = Nothing

End Function

Private Function cargaCreditosHH() As Boolean

    Dim cCreditos As New Collection
    Dim bCreditos As Boolean
    
    cargaCreditosHH = False
    
    Set cCreditos = obtenCreditosDesdeArchivoHH(bCreditos)
    
    If bCreditos = True Then
        
        fnLimpiaGrid sprCreditoNuevo
        
        Call fnLlenaTablaCollection(sprCreditoNuevo, cCreditos)
        
        'Call fnIdentificaCreditosNoAutorizados
        
        'sprCreditoNuevo.SetFocus
        
        cargaCreditosHH = True
        
    End If
    
End Function

Private Function obtenCreditosDesdeArchivoHH(ByRef bCreditos As Boolean) As Collection
    
    On Error GoTo ErrArchivo
    
    Dim strArchivo As String
    Dim iPreciosArchivo As Integer
    
    bCreditos = False

    strArchivo = dameArchivo()
    
    'Poner el cursor de proceso
    Screen.MousePointer = vbHourglass
    
    'LLevar los códigos a la base de datos
    If abreArchivofn(strArchivo, iPreciosArchivo, PARA_LECTURA) Then

        Dim Registros As New Collection
        Dim Registro As Collection
        Dim oCampo As New Campo
        Dim strCampo As String
        Dim strRegistro As String
        
        Dim iPosicion As Integer
        Dim iPosicionComa As Integer
        Dim iCliente As Integer
        
        'Dim oCredito As New credito
        Dim oCliente As New Cliente
        
        Do While Not EOF(iPreciosArchivo)
            
            Set Registro = New Collection
            
            strRegistro = ""
            
            obtenRegistrofn iPreciosArchivo, strRegistro
            
            iPosicion = 1
            
            Do
                
                iPosicionComa = InStr(iPosicion, strRegistro, ",")
                
                If iPosicionComa > 0 Then
                
                    strCampo = Mid(strRegistro, iPosicion, iPosicionComa - iPosicion)
                    
                    Registro.Add oCampo.CreaCampo(adInteger, , , strCampo)
'Modificacion 08/10/2010
                    If iPosicion = 1 Then

                        iCliente = Val(strCampo)

                        'Estos datos son para complemento de la operación
                        If oCliente.informacionGeneral(iCliente) = True Then
                        'If oCliente.fnInformacion(iCliente, Date) = True Then

                            Registro.Add oCampo.CreaCampo(adInteger, , , oCliente.mstrNombre + " " + oCliente.mstrApPaterno)    'Cliente

                        End If

                    End If
'FIn de Modificacion 08/10/2010

                    iPosicion = iPosicionComa + 1
                    
                End If
                    
            Loop Until iPosicionComa = 0
            
            strCampo = Mid(strRegistro, iPosicion, Len(strRegistro) - (iPosicion - 1))
            Registro.Add oCampo.CreaCampo(adInteger, , , strCampo) 'Cobrador
                       
            Registros.Add Registro
            
            bCreditos = True
            
        Loop
        
        cierraArchivofn iPreciosArchivo
        
        Set obtenCreditosDesdeArchivoHH = Registros
                
        'Set oCredito = Nothing
        Set oCliente = Nothing
        
    End If
    
'    If gEncriptado <> "NO" Then
'
'        Dim fs, fil1
'        Set fs = CreateObject("Scripting.FileSystemObject")
'
'        'Dim fso As New FileSystemObject, fil1 ', fil2
'
'        Set fil1 = fs.GetFile(strArchivo)
'
'        ' Delete the files.
'        fil1.Delete
'
'    End If
    
        'Poner el cursor Normal
        Screen.MousePointer = vbDefault

ErrArchivo:
    Exit Function
    
End Function

Private Function dameArchivo() As String

'    If gEncriptado = "NO" Then

        cmdListaCreditos.Filter = "CreditosNuevos(*.txt)|*.txt"
        cmdListaCreditos.FileName = "CreditosNuevos"
        cmdListaCreditos.DialogTitle = "Importar Créditos Nuevos"
        cmdListaCreditos.ShowOpen

        dameArchivo = cmdListaCreditos.FileName

'    Else

'        dameArchivo = "c:\MAP\ARCHIVOS SALIDA\CreditosNuevos_admin.txt"

'    End If
    
End Function

'Private Function obtenCreditosDesdeArchivoHH(ByRef bCreditos As Boolean) As Collection
'
'    On Error GoTo ErrArchivo
'
'    Dim strArchivo As String
'    Dim iPreciosArchivo As Integer
'
'    bCreditos = False
'
'    cmdListaCreditos.Filter = "CreditosNuevos(*.txt)|*.txt"
'    cmdListaCreditos.FileName = "CreditosNuevos"
'    cmdListaCreditos.DialogTitle = "Importar Créditos Nuevos"
'    cmdListaCreditos.ShowOpen
'
'    strArchivo = cmdListaCreditos.FileName
'
'    'LLevar los códigos a la base de datos
'    If abreArchivofn(strArchivo, iPreciosArchivo, PARA_LECTURA) Then
'
'        Dim Registros As New Collection
'        Dim Registro As Collection
'        Dim oCampo As New Campo
'        Dim strCampo As String
'        Dim iCliente As Integer
'        Dim strRegistro As String
'        Dim fCantidad As Double
'
'        Dim iPosicion As Integer
'        Dim iPosicionComa As Integer
'
'        Dim oCliente As New Cliente
'        Dim oCredito As New credito
'
'        Do While Not EOF(iPreciosArchivo)
'
'            Set Registro = New Collection
'
'            strRegistro = ""
'
'            obtenRegistrofn iPreciosArchivo, strRegistro
'
'            iPosicion = 1
'
'            'Registro.Add oCampo.CreaCampo(adInteger, , , 0)
'
'            Do
'
'                iPosicionComa = InStr(iPosicion, strRegistro, ",")
'
'                If iPosicionComa > 0 Then
'
'                    strCampo = Mid(strRegistro, iPosicion, iPosicionComa - iPosicion)
'
'                    Registro.Add oCampo.CreaCampo(adInteger, , , strCampo)
'
'                    If iPosicion = 1 Then
'
'                        iCliente = Val(strCampo)
'
'                        'Estos datos son para complemento de la operación
'                        If oCliente.fnInformacion(iCliente, Date) = True Then
'
'                            Registro.Add oCampo.CreaCampo(adInteger, , , oCliente.mstrNombre + " " + oCliente.mstrApPaterno)    'Cliente
'
'                        End If
'
'                    End If
'
'                    iPosicion = iPosicionComa + 1
'
'                End If
'
'            Loop Until iPosicionComa = 0
'
'            strCampo = Mid(strRegistro, iPosicion, Len(strRegistro) - (iPosicion - 1))
'            Registro.Add oCampo.CreaCampo(adInteger, , , strCampo) 'Cantidad
'
'            fCantidad = Val(strCampo)
'
'            Registro.Add oCampo.CreaCampo(adInteger, , , 14) 'Interes
'            Registro.Add oCampo.CreaCampo(adInteger, , , fCantidad * (14 / 100)) 'Financiamiento
'            Registro.Add oCampo.CreaCampo(adInteger, , , 30) 'No. de Pagos
'            Registro.Add oCampo.CreaCampo(adInteger, , , (fCantidad + (fCantidad * (14 / 100))) / 30) 'Cantidad a Pagar
'            Registro.Add oCampo.CreaCampo(adInteger, , , fCantidad + (fCantidad * (14 / 100))) 'Cantidad Total
'
'            If True = oCredito.validaDisponibilidadDeCredito(iCliente, fCantidad) Then
'                Registro.Add oCampo.CreaCampo(adInteger, , , 1) 'Disponibilidad de crédito
'            Else
'                Registro.Add oCampo.CreaCampo(adInteger, , , 0) 'No disponibilidad de crédito
'            End If
'
'            Registros.Add Registro
'
'            bCreditos = True
'
'        Loop
'
'        Set obtenCreditosDesdeArchivoHH = Registros
'
'        cierraArchivofn iPreciosArchivo
'
'        Set oCliente = Nothing
'        Set oCredito = Nothing
'
'    End If
'
'ErrArchivo:
'    Exit Function
'
'End Function

Private Function fnIdentificaCreditosNoAutorizados()

    Dim lRow As Long
    
    
    For lRow = 1 To sprCreditoNuevo.DataRowCnt
    
        sprCreditoNuevo.Col = COL_CREDITOS_AUTORIZADO
        sprCreditoNuevo.Row = lRow
        If Val(sprCreditoNuevo.Text) = 0 Then
            
            ' Lock block of cells
            ' Specify the block of cells
            sprCreditoNuevo.Col = 1
            sprCreditoNuevo.Col2 = -1
            sprCreditoNuevo.Row = lRow
            sprCreditoNuevo.Row2 = lRow
            ' Lock cells
            sprCreditoNuevo.BlockMode = True
            
            sprCreditoNuevo.BackColor = 8454143 'RGB(255, 0, 0)
            
            sprCreditoNuevo.BlockMode = False
            
        Else
        
            ' Lock block of cells
            ' Specify the block of cells
            sprCreditoNuevo.Col = 1
            sprCreditoNuevo.Col2 = -1
            sprCreditoNuevo.Row = lRow
            sprCreditoNuevo.Row2 = lRow
            ' Lock cells
            sprCreditoNuevo.BlockMode = True
            
            sprCreditoNuevo.BackColor = 12648384 'RGB(0, 0, 0)
            
            sprCreditoNuevo.BlockMode = False
            
        End If
        
    Next lRow
    
End Function

Private Sub sprCreditoNuevo_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case vbKeyDelete
            sprCreditoNuevo.Col = sprCreditoNuevo.ActiveCol
            sprCreditoNuevo.Row = sprCreditoNuevo.ActiveRow
            sprCreditoNuevo.Action = ActionDeleteRow
        
        Case vbKeyF7
            sprCreditoNuevo.Col = sprCreditoNuevo.ActiveCol
            sprCreditoNuevo.Row = sprCreditoNuevo.ActiveRow
            sprCreditoNuevo.Action = ActionDeleteRow
    
    End Select
    
End Sub

Private Sub sprCreditoNuevo_DblClick(ByVal Col As Long, ByVal Row As Long)

    sprCreditoNuevo.Col = COL_CREDITOS_CLIENTE
    sprCreditoNuevo.Row = Row
    
    If despliegaInformacionCliente(Val(sprCreditoNuevo.Text), Format(Now(), "dd/mm/yyyy")) = False Then
    
        'credito = 0
        MsgBox "No existen clientes con ese número ", vbInformation, "Consulta de Clientes"
        txtNombre.Text = ""
        txtNombre.SetFocus
        
    End If
    
    Select Case giTipoUsuario
        Case Is = USUARIO_GERENTE
            
            If sprCredito.DataRowCnt > 0 Then
                cmdpagos.Enabled = True
                cmdModificaCredito.Enabled = True
            Else
                cmdpagos.Enabled = False
                cmdModificaCredito.Enabled = False
            End If
        
        'Case Is = USUARIO_ADMINISTRADOR
        'Case Is = USUARIO_USUARIO
    
    End Select
        
End Sub

Private Sub cmdSaldosIndividuales_Click()
    
    sprCredito.Col = 1 'COL_FOLIO
    sprCredito.Row = sprCredito.ActiveRow
    
    If Val(sprCredito.Text) <= 0 Then
        MsgBox "Seleccione un cliente y su crédito", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    'sprCredito.Col = 1
    'sprCredito.Row = sprCredito.ActiveRow
        
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
        
    'Me.cmdSaldosIndividuales.SetFocus
    
End Sub

Private Sub cmdgrabar_Click()
            
    Dim oCliente As New Cliente
    
    If validaForma(False) = False Then
        
        Set oCliente = Nothing
        Exit Sub
        
    End If
        
    If cmdgrabar.Caption = "Alta" Then
        
        Dim iCliente As Integer
        iCliente = oCliente.altaCliente(obtenDatos(False))
        
        If iCliente > 0 Then
                            
            MsgBox "El cliente ha sido registrado", vbInformation, "Alta de Clientes"
            'nocliente = numero
            cmdgrabar.Caption = "Modificar"
        
            'txtapellido_LostFocus
            cmdcredito.Enabled = True
            cmdpagos.Enabled = True
            'limite = txtmaxcredito.Text
                    
            'Actualiza la lista de clientes
            If oCliente.listaLimiteCredito Then
                Call fnLlenaTablaCollection(sprCliente, oCliente.cDatos)
            End If
            
            buscaEnLista sprCliente, COL_CLIENTE_ID, iCliente
            
            txtnocliente.Text = CStr(iCliente)
            
        Else
            MsgBox "El cliente NO ha sido registrado", vbInformation, "Alta de Clientes"
        End If
        
    Else
            
        'Actualiza datos del cliente
        oCliente.cDetalle = obtenDatos(True)
        oCliente.actualizaCliente
        
        'base.Execute "update clientes set atrasospermitidos = " + IIf(txtAtrasos.Text = "", "2", txtAtrasos.Text) + ",nombre='" + txtnombre.Text + "',apellido='" + txtapellido.Text + "',direccion='" + txtdireccion.Text + "',ubicacion ='" + IIf(txtubicacion.Text = "", "Nulo", txtubicacion.Text) + "',cp='" + IIf(txtcp.Text = "", "Nulo", CStr(txtcp.Text)) + "',ciudad='" + IIf(txtcd.Text = "", "Nulo", txtcd.Text) + "',estado='" + IIf(txtedo.Text = "", "Nulo", txtedo.Text) + "',telefono='" + IIf(txttelefono.Text = "", "Nulo", CStr(txttelefono.Text)) + "',maxcredito='" + CStr(txtmaxcredito.Text) + "' where no_cliente=" + CStr(txtnocliente.Text)
        MsgBox "Los datos del cliente han sido modificados", vbInformation, "Modificación de Clientes"
        'cte = txtnocliente.Text
        sprCredito.Action = ActionClearText
        
        'txtnocliente.Text = cte
        
        'limite = txtmaxcredito.Text
        'nocliente = txtnocliente.Text
        cmdcredito.Enabled = True
        cmdpagos.Enabled = True
        
        'Actualiza la lista de clientes
        If oCliente.listaLimiteCredito Then
            Call fnLlenaTablaCollection(sprCliente, oCliente.cDatos)
        End If

    End If
    
    Set oCliente = Nothing
        
End Sub

Private Sub cmdNuevo_Click()
    
    txtnocliente.Text = ""
    txtnocliente.Enabled = True
    txtNombre.Text = ""
    txtapellido.Text = ""
    txtdireccion.Text = ""
    txtubicacion.Text = ""
    txtcd.Text = ""
    txtcp.Text = ""
    cbCobrador.ListIndex = -1
    txtTelefono.Text = ""
    txtmaxcredito.Text = 0
    txtAtrasos.Text = 2
    txtApellidoMaterno.Text = ""
    txtColonia.Text = ""
    txtEntidadFederativa.Text = ""
    cmdcredito.Enabled = False
    cmdpagos.Enabled = False
    cmdgrabar.Caption = "Alta"
    txtNombre.SetFocus
    
    fnLimpiaGrid sprCredito

End Sub

Private Sub cmdcredito_Click()
    
    Dim bAltaCredito As Boolean
    Dim iNoFactura As Long
    Dim iPagos As Integer
    
    creditofrm.strCobrador = cbCobrador.Text
    creditofrm.iNoCliente = Val(Me.txtnocliente.Text)
    creditofrm.strNombreCliente = Me.txtNombre & " " & Me.txtapellido
    creditofrm.iNoFactura = 0 'es un credito nuevo
    creditofrm.Show vbModal

    bAltaCredito = creditofrm.bAltaCredito
    iNoFactura = creditofrm.iNoFactura
    iPagos = creditofrm.iPagos
    
    If bAltaCredito = True Then
        
        If vbYes = MsgBox("¿Desea imprimir la Póliza?", vbQuestion + vbYesNo) Then
        
            Call imprimefn(iNoFactura, iPagos)
            
        End If
        
        'Actualiza la lista de creditos del cliente
        Dim oCliente As New Cliente
        Call fnLimpiaGrid(sprCredito)
        oCliente.fnInformacion Val(Me.txtnocliente.Text), Format(Now, "dd/mm/yyyy")
        If oCliente.creditosPorEstatus(Val(txtnocliente.Text), Format(Now(), "dd/mm/yyyy"), "V") = True Then
            Call fnLlenaTablaCollection(sprCredito, oCliente.cDatos)
        End If
        
        'Dim oCliente As New Cliente
        'oCliente.fnInformacion Val(Me.txtnocliente.Text), Format(Now, "dd/mm/yyyy")
        'tabCredito.SelectedItem.Index = TAB_VIGENTES
        'Call fnLlenaTablaCollection(sprCredito, oCliente.cCreditos(tabCredito.SelectedItem.Index))
        Set oCliente = Nothing
    End If
    
    'Me.cmdcredito.SetFocus
    
End Sub

Private Sub cmdModificaCredito_Click()
    
    sprCredito.Row = sprCredito.ActiveRow
    sprCredito.Col = CREDITO_FACTURA
    'If Len(sprCredito.Text) > 0 Then
        modificaCreditofrm.iNoFactura = Val(sprCredito.Text)
        modificaCreditofrm.iNoCliente = Val(txtnocliente.Text)
        modificaCreditofrm.strNombreCliente = txtNombre & " " & txtapellido
        modificaCreditofrm.bModifica = True
        modificaCreditofrm.Show vbModal
        Me.cmdModificaCredito.SetFocus
    'End If
    
    iRowActual_creditos = sprCredito.ActiveRow
'    sprCredito_DblClick CREDITO_FACTURA, sprCredito.ActiveRow
End Sub

Private Sub cmdpagos_Click()

    'accesofrm.Show vbModal

    'If accesofrm.bPermiteAcceso = True Then
    
        sprCredito.Col = 1 'COL_FOLIO
        sprCredito.Row = sprCredito.ActiveRow
        
        modificaPagofrm.lFolio = Val(sprCredito.Text)
        
        modificaPagofrm.Show vbModal
    
        Me.cmdpagos.SetFocus
    'End If
    
End Sub

Private Sub cmdsalir_Click()
    
    sicPrincipalfrm.Caption = ""
    sicPrincipalfrm.pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
    despliegaVentana portadafrm, WND_PORTADA
    
End Sub

Private Sub sprCliente_DblClick(ByVal Col As Long, ByVal Row As Long)

    sprCliente.Col = 1
    sprCliente.Row = Row
    
    If despliegaInformacionCliente(Val(sprCliente.Text), Format(Now(), "dd/mm/yyyy")) = False Then
    
        'credito = 0
        MsgBox "No existen clientes con ese número ", vbInformation, "Consulta de Clientes"
        txtNombre.Text = ""
        txtNombre.SetFocus
        
    End If
    
    Select Case giTipoUsuario
        Case Is = USUARIO_GERENTE
            
            If Val(Me.txtnocliente.Text) > 0 Then
                cmdcredito.Enabled = True
            Else
                cmdcredito.Enabled = False
            End If
            
            If sprCredito.DataRowCnt > 0 Then
                cmdpagos.Enabled = True
                cmdModificaCredito.Enabled = True
                cmdSaldosIndividuales.Enabled = True
                cmdImprimePoliza.Enabled = True
            'Else
            '    cmdpagos.Enabled = False
            '    cmdModificaCredito.Enabled = False
            '    cmdSaldosIndividuales.Enabled = False
            '    cmdImprimePoliza.Enabled = False
            End If
        
        'Case Is = USUARIO_ADMINISTRADOR
        'Case Is = USUARIO_USUARIO
    
    End Select
        
End Sub

Private Sub sprCliente_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    If Row <> NewRow Then
        If NewRow > 0 Then
                sprCliente.Col = 1
                sprCliente.Row = NewRow
        
                If despliegaInformacionCliente(Val(sprCliente.Text), Format(Now(), "dd/mm/yyyy")) = False Then
        
                    'credito = 0
                    MsgBox "No existen clientes con ese número ", vbInformation, "Consulta de Clientes"
                    txtNombre.Text = ""
                    txtNombre.SetFocus
        
                End If
        End If
    End If
End Sub

Private Function despliegaInformacionCliente(iCliente As Integer, strFecha As String) As Boolean

    Dim cCreditos As New Collection
    Dim oCliente As New Cliente

    Screen.MousePointer = vbHourglass

    If oCliente.fnInformacion(iCliente, strFecha) = True Then
    
    'If oCliente.cDatos.Count > 0 Then
        txtnocliente.Text = oCliente.iCliente
        txtNombre.Text = oCliente.mstrNombre
        txtapellido.Text = oCliente.mstrApPaterno
        txtApellidoMaterno.Text = oCliente.mstrApMaterno
        
'MsgBox "A Buscar texto en combo" & oCliente.mstrEstado, vbOKOnly
        fnBuscaTextoCombo cbCobrador, oCliente.mstrEstado
        
'MsgBox "A desplegar mas datos en pantalla", vbOKOnly
        txtdireccion.Text = oCliente.mstrDireccionCliente
        txtColonia.Text = oCliente.mstrColonia
        txtcd.Text = oCliente.mstrCiudad
        txtEntidadFederativa.Text = oCliente.mstrEntidadFederativa
        txtcp.Text = oCliente.mstrCP
        txtTelefono.Text = oCliente.mstrTelefonoCliente
        txtubicacion.Text = oCliente.mstrUbicacion
        
        txtmaxcredito.Text = Format(oCliente.mdCreditoMaximo, "$##,###")
        txtAtrasos.Text = oCliente.iAtrasosPermitidos
        
        Select Case giTipoUsuario
            Case Is = USUARIO_GERENTE
                
                cmdcredito.Enabled = True
                cmdpagos.Enabled = True
                cmdNuevo.Enabled = True
                cmdgrabar.Enabled = True
                cmdgrabar.Caption = "Modificar"
                cmdCargaCreditos.Enabled = True
                cmdImprimePoliza.Enabled = True
                cmdPagosUsuario.Enabled = True
                
            'Case Is = USUARIO_ADMINISTRADOR
            'Case Is = USUARIO_USUARIO
        
        End Select
        'nocliente = txtnocliente.Text
        'txtapellido.SetFocus
        paso = 0
        'txtnocliente.Enabled = False
        'credito = 1
    
        'Set cCreditos = oCliente.cCreditos
        Dim strStatus As String
        
        Select Case iTabActivo
            Case Is = TAB_VIGENTES
                strStatus = "V"
            Case Is = TAB_PENDIENTES
                strStatus = "P"
            Case Is = TAB_CANCELADOS
                strStatus = "C"
            Case Is = TAB_VENCIDOS
                strStatus = "E"
            Case Is = TAB_TERMINADOS
                strStatus = "T"
        End Select

'MsgBox "Voy por los créditos por estatus", vbOKOnly
        
        Call fnLimpiaGrid(sprCredito)
        If oCliente.creditosPorEstatus(iCliente, strFecha, strStatus) = True Then
            Call fnLlenaTablaCollection(sprCredito, oCliente.cDatos)
        End If
        
        'Call fnLlenaTablaCollection(sprCredito, cCreditos(iTabActivo))
    
        despliegaInformacionCliente = True
    Else
        despliegaInformacionCliente = False
    End If
    
    Screen.MousePointer = vbDefault
    
    Set oCliente = Nothing

End Function

Private Sub sprCredito_DblClick(ByVal Col As Long, ByVal Row As Long)

    sprCredito.Row = Row
    sprCredito.Col = CREDITO_FACTURA
    If Len(sprCredito.Text) > 0 Then
        modificaCreditofrm.iNoFactura = Val(sprCredito.Text)
        modificaCreditofrm.iNoCliente = Val(txtnocliente.Text)
        modificaCreditofrm.strNombreCliente = txtNombre & " " & txtapellido
        modificaCreditofrm.bModifica = False
        modificaCreditofrm.Show vbModal
    End If
    
    iRowActual_creditos = Row
    
End Sub

Private Sub sprCredito_GotFocus()
    iRowActual_creditos = sprCredito.ActiveRow
End Sub

Private Sub tabCredito_Click()

    If Val(txtnocliente.Text) <= 0 Then
        MsgBox "Seleccione el cliente por favor", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    If tabCredito.SelectedItem.Index = iTabActivo Then Exit Sub ' No need to change frame.

    Dim oCliente As New Cliente
    'oCliente.fnInformacion Val(txtnocliente.Text), Format(Now(), "dd/mm/yyyy")
    'Call fnLlenaTablaCollection(sprCredito, oCliente.cCreditos(tabCredito.SelectedItem.Index))
    'Call fnLlenaTablaCollection(sprCredito, cCreditos(tabCredito.SelectedItem.Index))
    
    Dim cCreditos As New Collection
    Dim strStatus As String
    
        Select Case tabCredito.SelectedItem.Index
            Case Is = TAB_VIGENTES
                strStatus = "V"
            Case Is = TAB_PENDIENTES
                strStatus = "P"
            Case Is = TAB_CANCELADOS
                strStatus = "C"
            Case Is = TAB_VENCIDOS
                strStatus = "E"
            Case Is = TAB_TERMINADOS
                strStatus = "T"
        End Select
        
        Call fnLimpiaGrid(sprCredito)
        If oCliente.creditosPorEstatus(Val(txtnocliente.Text), Format(Now(), "dd/mm/yyyy"), strStatus) = True Then
            Call fnLlenaTablaCollection(sprCredito, oCliente.cDatos)
        End If
    
    
    Select Case giTipoUsuario
        Case Is = USUARIO_GERENTE
            
            If sprCredito.DataRowCnt > 0 Then
                cmdpagos.Enabled = True
                cmdModificaCredito.Enabled = True
                cmdSaldosIndividuales.Enabled = True
            Else
                cmdpagos.Enabled = False
                cmdModificaCredito.Enabled = False
                cmdSaldosIndividuales.Enabled = False
            End If
        
        'Case Is = USUARIO_ADMINISTRADOR
        'Case Is = USUARIO_USUARIO
    
    End Select
        
    iTabActivo = tabCredito.SelectedItem.Index
    
    Set oCliente = Nothing

End Sub

Private Sub cmdVigentes_Click()

    'accesofrm.Show vbModal

    'If accesofrm.bPermiteAcceso = True Then

        Me.cmdVigentes.SetFocus
        Call fnImprime("clientesVigentes", crClientesVigentes, Format(Now(), "dd/mm/yyyy"), Format(Now(), "dd/mm/yyyy"))
    'End If
    
End Sub


Private Sub cmdPendientes_Click()
    'accesofrm.Show vbModal

    'If accesofrm.bPermiteAcceso = True Then

        Call fnImprime("clientesPendientes", crClientesPendientes, "", "")
        
    '    Me.cmdPendientes.SetFocus
        
    'End If
    
End Sub

Private Sub cmdVencidos_Click()

    'accesofrm.Show vbModal

    'If accesofrm.bPermiteAcceso = True Then
    
        periodofrm.Show vbModal
        
        Call fnImprime("ClientesVencidos", crClientesVencidos, periodofrm.strFechaInicial, periodofrm.strFechaFinal)
        
    '    Me.cmdVencidos.SetFocus
        
        
    'End If
    
End Sub

Private Sub cmdTerminados_Click()

    'accesofrm.Show vbModal

    'If accesofrm.bPermiteAcceso = True Then

        periodofrm.Show vbModal
        
        Call fnImprime("clientesTerminados", crClientesTerminados, periodofrm.strFechaInicial, periodofrm.strFechaFinal)
    
    '    Me.cmdTerminados.SetFocus
    'End If
    
End Sub

Private Sub cmdTotal_Click()
    
    'accesofrm.Show vbModal

    'If accesofrm.bPermiteAcceso = True Then

        periodofrm.Show vbModal
        
        Call fnImprime("clientesTotal", crClientesTotal, periodofrm.strFechaInicial, periodofrm.strFechaFinal)
    '    Me.cmdTotal.SetFocus
    'End If

End Sub

Private Sub cmdCancelados_Click()

    'accesofrm.Show vbModal

    'If accesofrm.bPermiteAcceso = True Then

        Call fnImprime("clientesCancelados", crClientesCancelados, "", "")
    '    Me.cmdCancelados.SetFocus
    'End If
    
End Sub

Private Sub fnImprime(strReporte As String, crObjeto As CrystalReport, _
                      strFechaInicial As String, strFechaFinal As String)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    Dim strNombreReporte As String
    'strFechaInicial = Date
    If strFechaInicial <> "" Then
        
'Cambio de regreso el 05/01/2012
'se hace efectivo de nuevo el periodo de consulta
        'cParametros.Add oCampo.CreaCampo(adInteger, , , Date)
        'cParametros.Add oCampo.CreaCampo(adInteger, , , Date)
    
        cParametros.Add oCampo.CreaCampo(adInteger, , , strFechaInicial)
        cParametros.Add oCampo.CreaCampo(adInteger, , , strFechaFinal)
'Cambio de regreso el 05/01/2012
        
    End If
    
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

'Private Sub grdclientes_DblClick()

'Dim datos, datos1 As Recordset
'Dim credito As Integer
'Dim fechas, total, pagado, abonado, fechaini As Double
'
'fechas = CDbl(CDate(Format(Now(), "dd/mm/yyyy")))
'
'grdclientes.Col = 0
''txtnocliente.Text = grdclientes.Text
'If grdclientes.Text <> "" Then
'    Set datos = base.OpenRecordset("select * from clientes where no_cliente=" & CStr(grdclientes.Text))
'    If datos.RecordCount > 0 Then
'
'        txtnocliente.Text = datos!no_cliente
'        txtnombre.Text = datos!Nombre
'        txtapellido.Text = datos!apellido
'        txtdireccion.Text = IIf(datos!DIRECCION = "Nulo", "", datos!DIRECCION)
'        txtubicacion.Text = IIf(datos!ubicacion = "Nulo", "", datos!ubicacion)
'        txtcd.Text = IIf(datos!ciudad = "Nulo", "", datos!ciudad)
'        txtedo.Text = IIf(datos!estado = "Nulo", "", datos!estado)
'        txtcp.Text = IIf(datos!cp = "Nulo", "", datos!cp)
'        txttelefono.Text = IIf(datos!TELEFONO = "Nulo", "", datos!TELEFONO)
'        txtmaxcredito.Text = Format(datos!maxcredito, "###,###,###,###0.00")
'        limite = datos!maxcredito
'        txtnocliente.Text = datos!no_cliente
'        txtAtrasos.Text = datos!atrasospermitidos
'
'        cmdcredito.Enabled = True
'        cmdpagos.Enabled = True
'        cmdnuevo.Enabled = True
'        cmdgrabar.Enabled = True
'        cmdgrabar.Caption = "Modificar"
'        nocliente = datos!no_cliente
'        credito = 1
'    Else
'        MsgBox "No existen clientes con ese número ", vbInformation, "Consulta de Clientes"
'        txtnombre.Text = ""
'        txtnombre.SetFocus
'    End If
'    datos.Close
'    If credito = 1 Then
'        grdcreditos.Clear
'        grdcreditos.Rows = 2
'        grdcreditos.Row = 0
'        grdcreditos.Col = 0
'        grdcreditos.Text = "Folio"
'        grdcreditos.Col = 1
'        grdcreditos.Text = "Monto"
'        grdcreditos.Col = 2
'        grdcreditos.Text = "Abonado"
'        grdcreditos.Col = 3
'        grdcreditos.Text = "Adeudo"
'        grdcreditos.Col = 4
'        grdcreditos.Text = "Status"
'        'grdcreditos.Col = 5
'        'grdcreditos.Text = "Fecha de Crédito"
'        grdcreditos.Col = 5
'        grdcreditos.Text = "Fecha de Inicio"
'        grdcreditos.Col = 6
'        grdcreditos.Text = "Fecha de Terminación"
'        grdcreditos.Col = 7
'        grdcreditos.Text = "Dias de Crédito"
'        grdcreditos.Col = 8
'        grdcreditos.Text = "Días Atraso"
'        grdcreditos.Col = 9
'        grdcreditos.Text = "Monto Atraso"
'
'        Set datos1 = base.OpenRecordset("select * from qrycreditoporcliente where no_cliente=" & CStr(txtnocliente.Text))
'        While Not datos1.EOF
'            If datos1!Status = "P" Or datos1!Status = "V" Then
'                grdcreditos.Rows = grdcreditos.Rows + 1
'                grdcreditos.Row = grdcreditos.Rows - 2
'                grdcreditos.Col = 0
'                grdcreditos.Text = datos1!factura
'                grdcreditos.Col = 1
'                grdcreditos.Text = Format(datos1!Canttotal, "###,###,###,###0.00")
'                total = CDbl(grdcreditos.Text)
'                grdcreditos.Col = 2
'                grdcreditos.Text = Format(IIf(IsNull(datos1!Cantpagada), 0, datos1!Cantpagada), "###,###,###,###0.00")
'                abonado = CDbl(grdcreditos.Text)
'                grdcreditos.Col = 3
'                grdcreditos.Text = Format(datos1!Canttotal - IIf(IsNull(datos1!Cantpagada), 0, datos1!Cantpagada), "###,###,###,###0.00")
'                adeudo = CDbl(grdcreditos.Text)
'                grdcreditos.Col = 4
'                grdcreditos.Text = datos1!Status
'                'grdcreditos.Col = 5
'                'grdcreditos.Text = datos1!fecha
'                grdcreditos.Col = 5
'                grdcreditos.Text = datos1!fechaini
'
'                '***************************************
'                fechaini = CDbl(datos1!fechaini)
'                '***************************************
'
'                grdcreditos.Col = 6
'                grdcreditos.Text = datos1!fechatermina
'                grdcreditos.Col = 7
'                If datos1!Status = "V" Or datos1!Status = "P" Then
'                    grdcreditos.Text = datos1!dias2
'                Else
'                    grdcreditos.Text = datos1!dias1
'                End If
'                grdcreditos.Col = 8
'                grdcreditos.Text = Format((IIf((fechas - fechaini) > datos1!no_pagos, datos1!no_pagos, (fechas - fechaini)) * (datos1!Canttotal / datos1!no_pagos) - CDbl(abonado)) / datos1!Cantpagar, "###,###,###,###0.00")
'                grdcreditos.Col = 9
'                grdcreditos.Text = Format((IIf((fechas - fechaini) > datos1!no_pagos, datos1!no_pagos, (fechas - fechaini)) * (datos1!Canttotal / datos1!no_pagos) - CDbl(abonado)), "###,###,###,###0.00")
'            End If
'            datos1.MoveNext
'        Wend
'        datos1.Close
'    End If
'End If
 
'End Sub

Private Sub txtapellido_KeyPress(KeyAscii As Integer)

    'Dim Nombre As String
    'Dim apellido As String
    'Dim datos As Recordset

    If KeyAscii = 13 Then
        If txtNombre.Text <> "" And txtapellido.Text <> "" Then
        
            cmdgrabar.Enabled = True
            cmdNuevo.Enabled = True
            cmdcredito.Enabled = False
            cmdpagos.Enabled = False
                
            Dim oCliente As New Cliente
            
            If oCliente.listaNameApellido(txtNombre.Text, txtapellido.Text) Then
                Call fnLlenaTablaCollection(sprCliente, oCliente.cDatos)
            End If
            
            Set oCliente = Nothing
            
        End If
                
'        Set datos = base.OpenRecordset("select * from clientes where Ucase(nombre)='" & UCase(txtnombre.Text) & "' and Ucase(apellido)='" & UCase(txtapellido.Text) & "'")
'        If datos.RecordCount > 0 Then
'            'MsgBox "El cliente ya existe", vbInformation, "Clientes"
'            txtnocliente.Text = datos!no_cliente
'            txtnocliente.Enabled = False
'            txtnombre.Text = datos!nombre
'            txtapellido.Text = datos!apellido
'            txtdireccion.Text = IIf(datos!direccion = "Nulo", "", datos!direccion)
'            txtubicacion.Text = IIf(datos!ubicacion = "Nulo", "", datos!ubicacion)
'            txtcd.Text = IIf(datos!ciudad = "Nulo", "", datos!ciudad)
'            txtcp.Text = IIf(datos!cp = "Nulo", "", datos!cp)
'            txtedo.Text = IIf(datos!estado = "Nulo", "", datos!estado)
'            txttelefono.Text = IIf(datos!telefono = "Nulo", "", datos!telefono)
'            txtmaxcredito.Text = Format(datos!maxcredito, "###,###,###,###0.00")
'            cmdgrabar.Caption = "Modificar"
'            cmdcredito.Enabled = True
'            cmdpagos.Enabled = True
'            limite = datos!maxcredito
'            nocliente = datos!no_cliente
'
'        Else
'            txtnocliente.Text = ""
'            txtnocliente.Enabled = False
'            txtdireccion.Text = ""
'            txtubicacion.Text = ""
'            txtcd.Text = ""
'            txtcp.Text = ""
'            txtedo.Text = ""
'            txttelefono.Text = ""
'            cmdgrabar.Caption = "Alta"
'            txtdireccion.SetFocus
'        End If
'        datos.Close
    End If

End Sub

Private Sub txtapellido_LostFocus()

'Call txtapellido_KeyPress(13)

End Sub


'Private Sub txtmaxcredito_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If txtmaxcredito.Text = "" Or Not IsNumeric(txtmaxcredito.Text) Then
'            txtmaxcredito.Text = 0
'        End If
'    End If
'End Sub

'Private Sub txtmaxcredito_LostFocus()
'    If txtmaxcredito.Text = "" Or Not IsNumeric(txtmaxcredito.Text) Then
'        txtmaxcredito.Text = 0
'    End If
'End Sub

Private Sub txtnocliente_KeyPress(KeyAscii As Integer)

'Dim datos As Recordset
'Dim datos1 As Recordset
    Dim credito As Integer
    Dim fechas, total, pagado, abonado, fechaini As Double
    
    fechas = CDbl(CDate(Format(Now(), "dd/mm/yyyy")))
    credito = 0

    If KeyAscii = 13 Then
        If Val(txtnocliente.Text) > 0 Then
        
            If despliegaInformacionCliente(Val(txtnocliente.Text), Format(Now(), "dd/mm/yyyy")) = False Then
                
                credito = 0
                MsgBox "No existen clientes con ese número ", vbInformation, "Consulta de Clientes"
                txtNombre.Text = ""
                txtNombre.SetFocus
            Else
                'busca el cliente en la lista y haz activa la linea
                buscaEnLista sprCliente, COL_CLIENTE_ID, Val(txtnocliente.Text)
                
                'Dim lRow As Long
                'sprCliente.Col = 1 'Id Cliente
                'For lRow = 1 To sprCliente.DataRowCnt
                '    sprCliente.Row = lRow
                '
                '    If Val(txtnocliente.Text) = Val(sprCliente.Text) Then
                '        sprCliente.Action = ActionActiveCell
                '        Exit For
                '    End If
                '
                'Next lRow
            End If
            
       
    '        Set datos = base.OpenRecordset("select * from clientes where no_cliente=" & CStr(txtnocliente.Text))
    '        If datos.RecordCount > 0 Then
    '
    '            txtnocliente.Text = datos!no_cliente
    '            txtnombre.Text = datos!Nombre
    '            txtapellido.Text = datos!apellido
    '            txtdireccion.Text = IIf(datos!DIRECCION = "Nulo", "", datos!DIRECCION)
    '            txtubicacion.Text = IIf(datos!ubicacion = "Nulo", "", datos!ubicacion)
    '            txtcd.Text = IIf(datos!ciudad = "Nulo", "", datos!ciudad)
    '            txtedo.Text = IIf(datos!estado = "Nulo", "", datos!estado)
    '            txtcp.Text = IIf(datos!cp = "Nulo", "", datos!cp)
    '            txttelefono.Text = IIf(datos!TELEFONO = "Nulo", "", datos!TELEFONO)
    '            txtmaxcredito.Text = Format(datos!maxcredito, "###,###,###,###0.00")
    '            limite = datos!maxcredito
    '            'txtnocliente.Text = datos!no_cliente
    '            txtAtrasos.Text = datos!atrasospermitidos
    '
    '            cmdcredito.Enabled = True
    '            cmdpagos.Enabled = True
    '            cmdnuevo.Enabled = True
    '            cmdgrabar.Enabled = True
    '            cmdgrabar.Caption = "Modificar"
    '            nocliente = datos!no_cliente
    '            txtapellido.SetFocus
    '            paso = 0
    '            txtnocliente.Enabled = False
    '            credito = 1
    '        Else
    '            credito = 0
    '            MsgBox "No existen clientes con ese número ", vbInformation, "Consulta de Clientes"
    '            txtnombre.Text = ""
    '            txtnombre.SetFocus
    '        End If
    '        datos.Close
    '        If credito = 1 Then
    '            grdcreditos.Clear
    '            grdcreditos.Rows = 2
    '            grdcreditos.Row = 0
    '            grdcreditos.Col = 0
    '            grdcreditos.Text = "Folio"
    '            grdcreditos.Col = 1
    '            grdcreditos.Text = "Monto"
    '            grdcreditos.Col = 2
    '            grdcreditos.Text = "Abonado"
    '            grdcreditos.Col = 3
    '            grdcreditos.Text = "Adeudo"
    '            grdcreditos.Col = 4
    '            grdcreditos.Text = "Status"
    '            'grdcreditos.Col = 5
    '            'grdcreditos.Text = "Fecha de Crédito"
    '            grdcreditos.Col = 5
    '            grdcreditos.Text = "Fecha de Inicio"
    '            grdcreditos.Col = 6
    '            grdcreditos.Text = "Fecha de Terminación"
    '            grdcreditos.Col = 7
    '            grdcreditos.Text = "Dias de Crédito"
    '            grdcreditos.Col = 8
    '            grdcreditos.Text = "Días Atraso"
    '            grdcreditos.Col = 9
    '            grdcreditos.Text = "Monto Atraso"
    '
    '            Set datos1 = base.OpenRecordset("select * from qrycreditoporcliente where no_cliente=" & CStr(txtnocliente.Text))
    '            While Not datos1.EOF
    '                If datos1!Status = "P" Or datos1!Status = "V" Then
    '                    grdcreditos.Rows = grdcreditos.Rows + 1
    '                    grdcreditos.Row = grdcreditos.Rows - 2
    '                    grdcreditos.Col = 0
    '                    grdcreditos.Text = datos1!factura
    '                    grdcreditos.Col = 1
    '                    grdcreditos.Text = Format(datos1!Canttotal, "###,###,###,###0.00")
    '                    total = CDbl(grdcreditos.Text)
    '                    grdcreditos.Col = 2
    '                    grdcreditos.Text = Format(IIf(IsNull(datos1!Cantpagada), 0, datos1!Cantpagada), "###,###,###,###0.00")
    '                    abonado = CDbl(grdcreditos.Text)
    '                    grdcreditos.Col = 3
    '                    grdcreditos.Text = Format(datos1!Canttotal - IIf(IsNull(datos1!Cantpagada), 0, datos1!Cantpagada), "###,###,###,###0.00")
    '                    adeudo = CDbl(grdcreditos.Text)
    '                    grdcreditos.Col = 4
    '                    grdcreditos.Text = datos1!Status
    '                    'grdcreditos.Col = 5
    '                    'grdcreditos.Text = datos1!fecha
    '                    grdcreditos.Col = 5
    '                    grdcreditos.Text = datos1!fechaini
    '
    '                    '***************************************
    '                    fechaini = CDbl(datos1!fechaini)
    '                    '***************************************
    '
    '                    grdcreditos.Col = 6
    '                    grdcreditos.Text = datos1!fechatermina
    '                    grdcreditos.Col = 7
    '                    If datos1!Status = "V" Or datos1!Status = "P" Then
    '                        grdcreditos.Text = datos1!dias2
    '                    Else
    '                        grdcreditos.Text = datos1!dias1
    '                    End If
    '                    grdcreditos.Col = 8
    '                    'IIf([cantpagar]=0,0,CDbl(Format((IIf((CDate(Format(Now(),"dd/mm/yyyy"))-[Fechaini]+1)>[no_pagos],[no_pagos],CDate(Format(Now(),"dd/mm/yyyy"))-[Fechaini]+1)*[cantpagar]-([creditos]![canttotal]-IIf(IsNull([qryclientescreditos]![Cantadeudada])=-1,[creditos]![Canttotal],[qryclientescreditos]![Cantadeudada])))/[Cantpagar],"Standard")))
    '                    grdcreditos.Text = Format((IIf((fechas - fechaini) > datos1!no_pagos, datos1!no_pagos, (fechas - fechaini)) * (datos1!Canttotal / datos1!no_pagos) - CDbl(abonado)) / datos1!Cantpagar, "###,###,###,###0.00")
    '
    '                    grdcreditos.Col = 9
    '                    grdcreditos.Text = Format((IIf((fechas - fechaini) > datos1!no_pagos, datos1!no_pagos, (fechas - fechaini)) * (datos1!Canttotal / datos1!no_pagos) - CDbl(abonado)), "###,###,###,###0.00")
    '
    '                End If
    '                datos1.MoveNext
    '            Wend
    '            datos1.Close
    '        End If
        End If
        
    End If
    
End Sub

'Private Sub txtnocliente_LostFocus()
'    Call txtnocliente_KeyPress(13)
'End Sub

'Private Sub txtnombre_KeyUp(KeyCode As Integer, Shift As Integer)

Private Sub txtnombre_KeyPress(KeyAscii As Integer)

'Dim datos As Recordset

    Dim oCliente As New Cliente
    If KeyAscii = 13 Then
    'If KeyCode = vbKeyReturn Then
    'If paso = 1 Then
        If txtNombre.Text <> "" Then
            If txtNombre.Text = "*" Then
                
                If oCliente.listaLimiteCredito Then
                    Call fnLlenaTablaCollection(sprCliente, oCliente.cDatos)
                End If

                'Set datos = base.OpenRecordset("select * from clientes")
            Else
                If oCliente.listaLikeName(txtNombre.Text) Then
                    Call fnLlenaTablaCollection(sprCliente, oCliente.cDatos)
                End If
                'Set datos = base.OpenRecordset("select * from clientes where Ucase(nombre) like '" & UCase(txtnombre.Text) & "*'")
            End If
'            If datos.RecordCount > 0 Then
'                Nombre = Trim(txtnombre.Text)
'                band = 1
'                frmlistaclientes.Show 1
'                txtnocliente.Enabled = False
'            Else
'                MsgBox "No existen clientes con ese nombre ", vbInformation, "Consulta de Clientes"
'                txtnombre.Text = ""
'                txtnombre.SetFocus
'            End If
'            datos.Close
        End If
    'End If
    End If
End Sub

