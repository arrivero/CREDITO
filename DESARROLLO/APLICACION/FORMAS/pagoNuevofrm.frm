VERSION 5.00
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form pagoNuevofrm 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "Pagos"
   ClientHeight    =   6480
   ClientLeft      =   315
   ClientTop       =   30
   ClientWidth     =   12540
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txttotadeudo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   6390
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H80000003&
      Caption         =   "Salir"
      Height          =   495
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5850
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H80000003&
      Caption         =   "Imprimir Pagos"
      Height          =   495
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5850
      Width           =   1335
   End
   Begin Crystal.CrystalReport crPagos 
      Left            =   6210
      Top             =   5910
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame frmPago 
      BackColor       =   &H80000003&
      Height          =   4785
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   12495
      Begin FPSpread.vaSpread sprPago 
         Height          =   4605
         Index           =   1
         Left            =   30
         TabIndex        =   0
         Top             =   150
         Width           =   12375
         _Version        =   196608
         _ExtentX        =   21828
         _ExtentY        =   8123
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15903391
         GridColor       =   15700869
         MaxCols         =   11
         ShadowColor     =   16038835
         SpreadDesigner  =   "pagoNuevofrm.frx":0000
      End
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   4950
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196608
      MinFontSize     =   1
      MaxFontSize     =   100
   End
   Begin VB.Frame frmPago 
      Height          =   4785
      Index           =   0
      Left            =   0
      TabIndex        =   35
      Top             =   570
      Width           =   12495
      Begin FPSpread.vaSpread sprPago 
         Height          =   2775
         Index           =   0
         Left            =   60
         TabIndex        =   36
         Top             =   1950
         Width           =   12375
         _Version        =   196608
         _ExtentX        =   21828
         _ExtentY        =   4895
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
         GrayAreaBackColor=   16777215
         MaxCols         =   10
         SpreadDesigner  =   "pagoNuevofrm.frx":1C3E
      End
   End
   Begin VB.CommandButton cmdregistra 
      BackColor       =   &H80000003&
      Caption         =   "Registra Pagos"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9780
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5850
      Width           =   1335
   End
   Begin VB.CommandButton cmdgraba 
      BackColor       =   &H80000003&
      Caption         =   "Pagos desde HH"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   8430
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5850
      Width           =   1335
   End
   Begin Threed.SSPanel pnlF7 
      Height          =   405
      Left            =   60
      TabIndex        =   27
      Top             =   5370
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   714
      _Version        =   196608
      BackColor       =   -2147483645
      Caption         =   "F7 - Para borrar la linea activa."
      BevelOuter      =   0
      Alignment       =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
   End
   Begin EditLib.fpCurrency txttotadeudoold 
      Height          =   405
      Left            =   7380
      TabIndex        =   30
      Top             =   5370
      Visible         =   0   'False
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
      BackColor       =   -2147483645
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
   Begin MSComDlg.CommonDialog cmdListaPagos 
      Left            =   4380
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   10
   End
   Begin Crystal.CrystalReport rpagos 
      Left            =   3930
      Top             =   5460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\facturas\pagos.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fraCaptura 
      Height          =   4275
      Left            =   30
      TabIndex        =   2
      Top             =   690
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox txtdias 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdagregar 
         Caption         =   "Agregar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7800
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtadeudo 
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
         Height          =   405
         Left            =   6120
         TabIndex        =   15
         Top             =   480
         Width           =   2910
      End
      Begin VB.TextBox txtpago 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   14
         Top             =   480
         Width           =   2910
      End
      Begin VB.TextBox txtcte 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtnombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtfolio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   2910
      End
      Begin VB.Frame Frame3 
         Caption         =   "Pagos Atrasados"
         Height          =   1065
         Left            =   210
         TabIndex        =   6
         Top             =   1860
         Width           =   5085
         Begin VB.TextBox txtdias_atraso 
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
            Left            =   120
            TabIndex        =   8
            Top             =   450
            Width           =   1980
         End
         Begin VB.TextBox txtcantatraso 
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
            Left            =   3150
            TabIndex        =   7
            Top             =   450
            Width           =   1800
         End
         Begin VB.Label Label12 
            Caption         =   "Dias Atraso"
            Height          =   255
            Left            =   150
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Cantidad Atrasada"
            Height          =   255
            Left            =   3180
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtultimo 
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
         Left            =   1710
         TabIndex        =   5
         Top             =   3180
         Width           =   1335
      End
      Begin VB.TextBox txtpago1 
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
         Left            =   3990
         TabIndex        =   4
         Top             =   3180
         Width           =   1755
      End
      Begin VB.TextBox txtadeudo1 
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
         Left            =   7230
         TabIndex        =   3
         Top             =   3180
         Width           =   1785
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "D�as de Cr�dito:"
         Height          =   255
         Left            =   3480
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Adeudo"
         Height          =   255
         Left            =   6150
         TabIndex        =   25
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Pago"
         Height          =   255
         Left            =   3270
         TabIndex        =   24
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "No Cliente:"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1485
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Folio"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "�ltimo Folio:"
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
         Left            =   150
         TabIndex        =   20
         Top             =   3225
         Width           =   1575
      End
      Begin VB.Label Label9 
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
         Left            =   6150
         TabIndex        =   19
         Top             =   3225
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   3270
         TabIndex        =   18
         Top             =   3225
         Width           =   1575
      End
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   30
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   873
      _Version        =   196608
      ForeColor       =   49152
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
      Begin VB.Label txtfecha 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   39
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Label txttotpago 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
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
      Left            =   10080
      TabIndex        =   41
      Top             =   5400
      Width           =   2355
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9990
      TabIndex        =   34
      Top             =   135
      Width           =   735
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
      Height          =   345
      Left            =   8415
      TabIndex        =   32
      Top             =   5430
      Width           =   1515
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000003&
      Caption         =   "Total Adeudos:"
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
      Left            =   4500
      TabIndex        =   31
      Top             =   5430
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "pagoNuevofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adeudo, pagogen As Double
Dim bandera, renglon As Integer

Public iCliente As Integer
Public iFactura As Integer
Public iAutomaticoManual As Integer
Public strUsuario As String
Public strNombreUsuarioPagos As String

Private Const COL_FOLIO = 1
Private Const COL_NO_CLIENTE = 2
Private Const COL_NOMBRE_CLIENTE = 3
Private Const COL_DIAZ_CREDITO = 4
Private Const COL_FECHA = 5
Private Const COL_PAGO = 6
Private Const COL_ADEUDO = 7
Private Const COL_HORA = 8
Private Const COL_DIAS_ATRAZO = 9
Private Const COL_MONTO_ATRAZO = 10

Private Const MODO_AUTOMATICO = 1
Private Const MODO_MANUAL = 0

Private bGrabado As Boolean

Private bActiva As Boolean

Private Sub cmdsalir_Click()
    sicPrincipalfrm.Caption = ""
    sicPrincipalfrm.pnlTitulo.Caption = "Soluci�n Integral de Administraci�n de Creditos"
    despliegaVentana portadafrm, WND_PORTADA
End Sub

'Private Sub Form_Activate()
'
'    If bActiva = False Then
'        If iAutomaticoManual = 1 Then   'Trae los pagos registrados (de la hand held)
'            frmPago(MODO_AUTOMATICO).Visible = True
'            frmPago(MODO_MANUAL).Visible = False
'            pnlF7.Visible = False
'
'            Call despliegaPagosHH
'
'        Else                            'Prepara para captura manual
'            frmPago(MODO_AUTOMATICO).Visible = False
'            frmPago(MODO_MANUAL).Visible = True
'            pnlF7.Visible = True
'        End If
'
'        txtfecha.Text = Format(Now, "dd/mm/yyyy")
'
'        bActiva = True
'    End If
'
'    bGrabado = False
'
'End Sub

Private Sub Form_Load()
            
    If iAutomaticoManual = 1 Then   'Trae los pagos registrados (de la hand held)
        
       
        frmPago(MODO_AUTOMATICO).Visible = True
        frmPago(MODO_MANUAL).Visible = False
        pnlF7.Visible = False

        Call despliegaPagosHH
        
    Else                            'Prepara para captura manual
        frmPago(MODO_AUTOMATICO).Visible = False
        frmPago(MODO_MANUAL).Visible = True
        pnlF7.Visible = True
    End If

    txtfecha.Caption = Format(Now, "dd/mm/yyyy")
    
    bGrabado = False
        
End Sub

'Private Function cargaPagosHH() As Boolean
'
'    Dim cPagos As Collection
'    Dim bPagos As Boolean
'
'    grabaPagosHH = False
'
'    Set cPagos = obtenPagosDesdeArchivoHH(bPagos)
'
'    If bPagos = True Then
'
'        Call fnLlenaTablaCollection(sprPago(MODO_AUTOMATICO), oPago.cDatos)
'
'        txttotpago.Text = Format(obtenTotalGrid(sprPago(MODO_AUTOMATICO), COL_PAGO), "###,###,###,###0.00")
'        txttotadeudo.Text = Format(obtenTotalGrid(sprPago(MODO_AUTOMATICO), COL_ADEUDO), "###,###,###,###0.00")
'
'    End If
'
'End Function

Private Function grabaPagosHH() As Boolean

    Dim cPagos As Collection
    Dim bPagos As Boolean
    
    grabaPagosHH = False
    
    Set cPagos = obtenPagosDesdeArchivoHH(bPagos)
    
    If bPagos = True Then
        
        Dim oPago As New Pago
        
            grabaPagosHH = oPago.grabaPagos(cPagos)
            
        Set oPago = Nothing
        
    End If
    
End Function

Private Sub cmdImprimir_Click()

        fnImprime "rpPagos", crPagos, "alex"

End Sub

Private Function obtenPagosDesdeArchivoHH(ByRef bPagos As Boolean) As Collection
    
    On Error GoTo ErrArchivo
    
    Dim strArchivo As String
    Dim iPreciosArchivo As Integer
    
    bPagos = False
 
    strArchivo = dameArchivo()
        
    'LLevar los c�digos a la base de datos
    If abreArchivofn(strArchivo, iPreciosArchivo, PARA_LECTURA) Then

        Dim Registros As New Collection
        Dim Registro As Collection
        Dim oCampo As New Campo
        Dim strCampo As String
        Dim lOrden As Long
        Dim lFolio As Long
        Dim strRegistro As String
        
        Dim iPosicion As Integer
        Dim iPosicionComa As Integer
        
        Dim oCredito As New credito
        
        lOrden = 1
        Do While Not EOF(iPreciosArchivo)
            
            Set Registro = New Collection
            
            strRegistro = ""
            
            obtenRegistrofn iPreciosArchivo, strRegistro
            
            iPosicion = 1
            
            Do
                
                iPosicionComa = InStr(iPosicion, strRegistro, ",")
                
                If iPosicionComa > 0 Then
                
                    strCampo = Mid(strRegistro, iPosicion, iPosicionComa - iPosicion)
                    If iPosicion = 1 Then
                        lFolio = Val(strCampo)
                    End If
                    iPosicion = iPosicionComa + 1
                    Registro.Add oCampo.CreaCampo(adInteger, , , strCampo)
                    
                End If
                    
            Loop Until iPosicionComa = 0
            
            strCampo = Mid(strRegistro, iPosicion, Len(strRegistro) - (iPosicion - 1))
            Registro.Add oCampo.CreaCampo(adInteger, , , strCampo) 'Lugar
                       
            Registro.Add oCampo.CreaCampo(adInteger, , , lOrden) 'Orden
                       
            Registros.Add Registro
            
            lOrden = lOrden + 1
            
            bPagos = True
            
        Loop
        
        Set obtenPagosDesdeArchivoHH = Registros
        
        cierraArchivofn iPreciosArchivo
        
        Set oCredito = Nothing
        
    End If
    
    If gEncriptado <> "NO" Then
    
        Dim fs, fil1
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        'Dim fso As New FileSystemObject, fil1 ', fil2
        
        Set fil1 = fs.GetFile(strArchivo)
        
        ' Delete the files.
        fil1.Delete
    
    End If

ErrArchivo:
    Exit Function

End Function

Private Function dameArchivo() As String

    
    If gEncriptado = "NO" Then
        
        cmdListaPagos.Filter = "Pagos(*.txt)|*.txt"
        cmdListaPagos.FileName = "" '"Pagos"
        cmdListaPagos.DialogTitle = "Importar pagos de cr�ditos."
        cmdListaPagos.ShowOpen
        
        dameArchivo = cmdListaPagos.FileName
        
    Else
    
        'dameArchivo = "c:\MAP\ARCHIVOS SALIDA\Pagos_admin.txt"
        dameArchivo = "c:\MAP\ARCHIVOS SALIDA\Pagos_" + strNombreUsuarioPagos + ".txt"
        
    End If
    
    
End Function
    
'Public Sub obtenPagosHandHeld(strNombreArchivo As String)
'
'    If File1.FileName = "Pagos.txt" Then
'
'        Dim ruta As String
'        Dim textoleido As String
'        Dim coma As Integer
'        Dim dato As String
'        Dim tam As Integer
'        Dim adeudo As Long
'        Dim datos As Recordset
'
'        If File1.Path = "C:\" Then
'            ruta = File1.Path & File1.FileName
'        Else
'            ruta = File1.Path & "\" & File1.FileName
'        End If
'
'        Open (ruta) For Input As #1
'
'            Dim Columna As Integer
'            Dim renglon As Integer
'            Dim folio As Long
'            Dim pago As Long
'            Dim Cliente As Long
'
'            renglon = 1
'            Columna = 1
'            While Not EOF(1)
'                grdpagos.Row = renglon
'                Line Input #1, textoleido
'                tam = Len(textoleido)
'                grdpagos.Col = Columna
'                coma = InStr(textoleido, ",")
'                dato = Trim(Mid(textoleido, 1, coma - 1))
'                tam = tam - coma
'                textoleido = Mid(textoleido, coma + 1, tam)
'                grdpagos.Text = dato
'                Columna = Columna + 1
'                folio = CDbl(dato)
'                Set datos = base.OpenRecordset("select creditos.no_cliente,clientes.nombre,clientes.apellido from creditos,clientes where creditos.no_cliente = clientes.no_cliente and factura=" & CStr(folio))
'                If datos.RecordCount > 0 Then
'                    grdpagos.Col = Columna
'                    grdpagos.Text = datos!no_cliente
'                    Columna = Columna + 1
'                    grdpagos.Col = Columna
'                    grdpagos.Text = CStr(datos!Nombre) + " " + CStr(datos!apellido)
'                    Cliente = datos!no_cliente
'                    Columna = Columna + 1
'                    datos.Close
'                Else
'                    grdpagos.Col = Columna
'                    grdpagos.Text = 0
'                    Columna = Columna + 1
'                    datos.Close
'                End If
'                    While coma <> 0
'                    grdpagos.Col = Columna
'                    coma = InStr(textoleido, ",")
'                    If coma = 0 Then
'                        dato = Trim(textoleido)
'
'                    Else
'                        dato = Trim(Mid(textoleido, 1, coma - 1))
'                        tam = tam - coma
'                        textoleido = Mid(textoleido, coma + 1, tam)
'                    End If
'                        grdpagos.Text = dato
'                        Columna = Columna + 1
'                Wend
'                    Set datos = base.OpenRecordset("select * from pagado where no_cliente=" & CStr(Cliente) & " and factura=" & CStr(folio))
'                    If datos.RecordCount > 0 Then
'                    grdpagos.Col = 6
'                    pago = grdpagos.Text
'                    grdpagos.Col = Columna
'                    grdpagos.Text = datos!Canttotal - datos!Cantpagada - pago
'                    datos.Close
'                    Else
'                    grdpagos.Col = 5
'                    grdpagos.Col = Columna
'                    grdpagos.Text = 0
'                    datos.Close
'                    End If
'
'                renglon = renglon + 1
'                grdpagos.Rows = grdpagos.Rows + 1
'                Columna = 1
'            Wend
'            Close #1
'    Else
'        MsgBox "El archivo no es un archivo de Pagos", vbOKOnly, "Pagos"
'    End If
'
'End Sub

Private Function despliegaPagosHH() As Boolean

    Dim cPagos As Collection
    Dim bPagos As Boolean
    
    despliegaPagosHH = False
    
    Set cPagos = obtenPagosDesdeArchivoHH(bPagos)
    
    If bPagos = True Then
    
        Dim oPago As New Pago
        
        Screen.MousePointer = vbHourglass
                
        oPago.grabaPagosTmp cPagos
        
        oPago.obtenPagosTmp
            
        sprPago(MODO_AUTOMATICO).ReDraw = False
        fnLimpiaGrid sprPago(MODO_AUTOMATICO)
        
        Call fnLlenaTablaCollection(sprPago(MODO_AUTOMATICO), oPago.cDatos)
        sprPago(MODO_AUTOMATICO).ReDraw = True
    
        txttotpago.Caption = Format(obtenTotalGrid(sprPago(MODO_AUTOMATICO), COL_PAGO), "$#,####.00")
        txttotadeudo.Text = obtenTotalGrid(sprPago(MODO_AUTOMATICO), COL_ADEUDO)
        
        Screen.MousePointer = vbDefault
        
        cmdregistra.Enabled = True
        
        Set oPago = Nothing
    
    End If
        
End Function

Private Function despliegaDatosCredito(cCredito As Collection) As String

    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    Dim fCantidadTotal As Double
    Dim fCantidadPagada As Double
    Dim strNombreCliente As String
    Dim fechas, fechaini, pagado As Double
    Dim pagos As Integer
    Dim total, fAdeudo As Double
    Dim strMensaje As String
    
    despliegaDatosCredito = ""
    
    fechas = CDbl(CDate(Format(Now(), "dd/mm/yyyy")))
    
    Set cRegistro = cCredito(1)
    
    Set oCampo = cRegistro(10) 'Fecha
    sprPago(MODO_MANUAL).Row = sprPago(MODO_MANUAL).ActiveRow
    If CDbl(Now) < CDbl(oCampo.Valor) Then
        
        despliegaDatosCredito = "La fecha de inicio de pago es mayor a hoy"
        Exit Function
        
    End If
    
    Set oCampo = cRegistro(18) 'Cantidad total
    fCantidadTotal = oCampo.Valor
    Set oCampo = cRegistro(19) 'Cantidad total
    fCantidadPagada = oCampo.Valor
    
    Set oCampo = cRegistro(2) 'Folio del cr�dito
    If fCantidadTotal - fCantidadPagada <= 0 Then
        despliegaDatosCredito = "�La deuda correspondiente al documento " & oCampo.Valor & " ya esta liquidada!"
        Exit Function
    End If
    
    Set oCampo = cRegistro(1) 'CLiente
    sprPago(MODO_MANUAL).Col = COL_NO_CLIENTE
    sprPago(MODO_MANUAL).Text = oCampo.Valor
    
    Set oCampo = cRegistro(4) 'Cantidad a pagar
    sprPago(MODO_MANUAL).Col = COL_PAGO
    sprPago(MODO_MANUAL).Text = Format(oCampo.Valor, "###,###,###,###0.00")
    pagogen = Format(oCampo.Valor, "###,###,###,###0.00")
    
    Set oCampo = cRegistro(6) 'Cantidad total
    'adeudo = Format(oCampo.Valor, "###,###,###,###0.00")
    'sprPago(MODO_MANUAL).Col = COL_ADEUDO
    'sprPago(MODO_MANUAL).Text = Format(adeudo, "###,###,###,###0.00")
    total = oCampo.Valor
    
    Set oCampo = cRegistro(8) 'Fecha inicial del cr�dito
    fechaini = CDbl(oCampo.Valor)
    
    sprPago(MODO_MANUAL).Col = COL_FECHA 'Fecha hoy
    sprPago(MODO_MANUAL).Text = CDate(Format(Now, "dd/mm/yyyy"))
    
    Set oCampo = cRegistro(7) 'No de Pagos
    pagos = oCampo.Valor
    sprPago(MODO_MANUAL).Col = COL_DIAZ_CREDITO
    sprPago(MODO_MANUAL).Text = oCampo.Valor 'CDate(Format(Now, "dd/mm/yyyy")) - CDate(oCampo.Valor) - 1
    
    Set oCampo = cRegistro(20) 'Nombre cliente
    strNombreCliente = oCampo.Valor
    Set oCampo = cRegistro(21) 'Apellido cliente
    sprPago(MODO_MANUAL).Col = COL_NOMBRE_CLIENTE
    sprPago(MODO_MANUAL).Text = strNombreCliente + " " + oCampo.Valor
    
    sprPago(MODO_MANUAL).Col = COL_ADEUDO
    fAdeudo = fCantidadTotal - fCantidadPagada
    If fAdeudo > pagogen Then
        sprPago(MODO_MANUAL).Text = Format(fAdeudo - pagogen, "###,###,###,###0.00")
    Else
        sprPago(MODO_MANUAL).Text = Format(pagogen, "###,###,###,###0.00")
    End If
    pagado = fCantidadPagada
    
    sprPago(MODO_MANUAL).Col = COL_DIAS_ATRAZO
    sprPago(MODO_MANUAL).Text = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado) / pagogen, "###,###,###,###0.00")
    sprPago(MODO_MANUAL).Col = COL_MONTO_ATRAZO
    sprPago(MODO_MANUAL).Text = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado), "###,###,###,###0.00")

    txttotpago.Caption = Format(obtenTotalGrid(sprPago(MODO_MANUAL), COL_PAGO), "$#,####.00")
    txttotadeudo.Text = Format(obtenTotalGrid(sprPago(MODO_MANUAL), COL_ADEUDO), "$#,####.00")
   
End Function


Private Sub sprPago_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = MODO_AUTOMATICO Then
        Exit Sub
    End If
    
    If vbKeyReturn = KeyAscii Then
        Dim lFactura As Long
        Dim lRowAnt As Long
        Dim lColAnt As Long
        Dim strMsg As String
        
        Select Case sprPago(MODO_MANUAL).ActiveCol
        
            Case Is = COL_FOLIO
                
                sprPago(MODO_MANUAL).Row = sprPago(MODO_MANUAL).ActiveRow
                sprPago(MODO_MANUAL).Col = sprPago(MODO_MANUAL).ActiveCol
                lRowAnt = sprPago(MODO_MANUAL).ActiveRow
                lColAnt = sprPago(MODO_MANUAL).ActiveCol
                
                lFactura = CLng(sprPago(MODO_MANUAL).Text)
                
                If existeValorEnGrid(sprPago(MODO_MANUAL), COL_FOLIO, lFactura, lRowAnt) = True Then
                    
                    pnlMsg.Caption = "�El No. de Factura " & lFactura & " ya se captur�, verifique por favor!" ', vbInformation + vbOKOnly
                    sprPago(MODO_MANUAL).Col = COL_PAGO
                    sprPago(MODO_MANUAL).Row = lRowAnt
                    sprPago(MODO_MANUAL).Action = ActionActiveCell
                    sprPago(MODO_MANUAL).Row = lRowAnt
                    sprPago(MODO_MANUAL).Col = lColAnt
                    
                Else
                    
                    Dim oPago As New Pago
                    
                    If oPago.registrado(lFactura, Format(Now(), "dd/mm/yyyy")) Then
                        
                        pnlMsg.Caption = "�Ya hay un pago registrado el d�a de hoy para el cr�dito con folio" & " " & sprPago(MODO_MANUAL).Text & ", " & "verifique por favor!"
                        sprPago(MODO_MANUAL).Col = COL_PAGO
                        sprPago(MODO_MANUAL).Row = lRowAnt
                        sprPago(MODO_MANUAL).Action = ActionActiveCell
                        
                    Else
                    
                        Dim oCredito As New credito
                        
                        If oCredito.datosCredito(lFactura) = False Then
                        
                            pnlMsg.Caption = "�No hay un cr�dito con No. de folio " & lFactura & ", " & "verifique por favor!"
                            sprPago(MODO_MANUAL).Col = COL_PAGO
                            sprPago(MODO_MANUAL).Row = lRowAnt
                            sprPago(MODO_MANUAL).Action = ActionActiveCell
                            
                        Else
                        
                            
                            lRowAnt = sprPago(MODO_MANUAL).ActiveRow
                            lColAnt = sprPago(MODO_MANUAL).ActiveCol
                            
                            strMsg = despliegaDatosCredito(oCredito.cDatos)
                            
                            If Len(strMsg) > 0 Then
                                pnlMsg = strMsg
                            End If
                            sprPago(MODO_MANUAL).Col = COL_PAGO
                            sprPago(MODO_MANUAL).Row = lRowAnt
                            sprPago(MODO_MANUAL).Action = ActionActiveCell
                                                    
                        End If
                        
                        Set oCredito = Nothing
                        
                    End If
                    
                    Set oPago = Nothing
                    
                End If
                
            Case Is = COL_PAGO
                
                If bGrabado = True Then
                    
                    sprPago(MODO_MANUAL).Row = sprPago(MODO_MANUAL).ActiveRow
                    sprPago(MODO_MANUAL).Col = COL_HORA
                    sprPago(MODO_MANUAL).Text = Format(Now, "hh:mm AM/PM")
                Else
                
                    If Len(pnlMsg.Caption) > 0 Then
                        pnlMsg.Caption = ""
                        sprPago(MODO_MANUAL).Col = COL_FOLIO
                        sprPago(MODO_MANUAL).Row = sprPago(MODO_MANUAL).ActiveRow
                        sprPago(MODO_MANUAL).Action = ActionActiveCell
                        sprPago(MODO_MANUAL).Text = ""
                    Else
                        sprPago(MODO_MANUAL).Col = COL_FOLIO
                        sprPago(MODO_MANUAL).Row = sprPago(MODO_MANUAL).ActiveRow + 1
                        sprPago(MODO_MANUAL).Action = ActionActiveCell
                        'Actualiza totales
                        txttotpago.Caption = Format(obtenTotalGrid(sprPago(MODO_MANUAL), COL_PAGO), "$#,####.00")
                        txttotadeudo.Text = Format(obtenTotalGrid(sprPago(MODO_MANUAL), COL_ADEUDO), "$#,####.00")

                    End If
                
                End If
                
        End Select
        
        If sprPago(MODO_MANUAL).DataRowCnt > 0 Then
            cmdgraba.Enabled = True
        End If
        
    End If
        
End Sub

Private Sub sprPago_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If Index = MODO_AUTOMATICO Then
        Exit Sub
    End If
    
    Select Case KeyCode

          Case vbKeyDelete, vbKeyF7

            sprPago(MODO_MANUAL).Row = sprPago(MODO_MANUAL).ActiveRow
            sprPago(MODO_MANUAL).Action = ActionDeleteRow
            'Actualiza totales
            txttotpago.Caption = Format(obtenTotalGrid(sprPago(MODO_MANUAL), COL_PAGO), "$,####.00")
            txttotadeudo.Text = Format(obtenTotalGrid(sprPago(MODO_MANUAL), COL_ADEUDO), "$#,####.00")

    End Select

End Sub

Private Function obtenPagos(bRegistraGraba As Boolean) As Collection

    Dim cPagos As New Collection
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    Dim strUsuarioTemp As String
    Dim strHora As String
    Dim strLugar As String
    Dim iRegistraGraba As Integer
    Dim lRow As Long
    
    For lRow = 1 To sprPago(MODO_MANUAL).DataRowCnt
    
        sprPago(MODO_MANUAL).Row = lRow
        
        If bRegistraGraba = False Then
            strUsuarioTemp = "Nombre"
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , "Nombre")
            strHora = ""
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , "") 'Hora
            strLugar = ""
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , "") 'Lugar
            iRegistraGraba = 0
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , 0) 'Registra
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'Consecutivo
        Else
            strUsuarioTemp = strUsuario
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , strUsuario)
            sprPago(MODO_MANUAL).Col = COL_HORA
            strHora = sprPago(MODO_MANUAL).Text
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago(MODO_MANUAL).Text)
            strLugar = "captura"
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , "captura") 'Lugar
            iRegistraGraba = 1
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'Graba
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'Consecutivo
        End If
        
        sprPago(MODO_MANUAL).Col = COL_FOLIO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(sprPago(MODO_MANUAL).Text))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , 0) 'No. de Pago
        cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'Consecutivo
        sprPago(MODO_MANUAL).Col = COL_PAGO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPago(MODO_MANUAL).Text)))
        cRegistro.Add oCampo.CreaCampo(adInteger, , , strUsuarioTemp) 'Usuario
        sprPago(MODO_MANUAL).Col = COL_FECHA
        cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago(MODO_MANUAL).Text)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , strHora) 'Hora
        cRegistro.Add oCampo.CreaCampo(adInteger, , , strLugar) 'Lugar
        
        cRegistro.Add oCampo.CreaCampo(adInteger, , , lRow) 'Orden
        cRegistro.Add oCampo.CreaCampo(adInteger, , , iRegistraGraba) 'Graba = 1, Registra = 0
        
        sprPago(MODO_MANUAL).Col = COL_NO_CLIENTE
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(sprPago(MODO_MANUAL).Text))
        sprPago(MODO_MANUAL).Col = COL_ADEUDO
        cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(sprPago(MODO_MANUAL).Text)))
        
'        If bRegistraGraba = False Then
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , "Nombre")
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , 0) 'Registra
'
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'Consecutivo
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , "") 'Hora
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , "") 'Lugar
'        Else
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , strUsuario)
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'Graba
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , 1) 'Consecutivo
'            sprPago(MODO_MANUAL).Col = COL_HORA
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , sprPago(MODO_MANUAL).Text)
'            cRegistro.Add oCampo.CreaCampo(adInteger, , , "captura") 'Lugar
'        End If
        
        cPagos.Add cRegistro
        
    Next lRow
    
    Set obtenPagos = cPagos
    
End Function

Private Sub fnImprime(strReporte As String, crObjeto As CrystalReport, strUsuario As String)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    Dim strNombreReporte As String
    
    'If strFechaInicial <> "" Then
    '
        cParametros.Add oCampo.CreaCampo(adInteger, , , strUsuario)
    '    cParametros.Add oCampo.CreaCampo(adInteger, , , strFechaFinal)
    '
    'End If
    
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

    'accesofrm.Show vbModal

    'sicPrincipalfrm.pnlTitulo.Caption = "Pagos - (" & UCase(gstrUsuario) & ")"
    
    'If accesofrm.bPermiteAcceso = True Then

    '    Dim strUsuario As String
    '    strUsuario = gstrUsuario

        If vbCancel = MsgBox("Coloque su equipo en la base y descargue los pagos, al terminar haga clic en 'Ok' para registrar sus pagos", vbInformation + vbOKCancel) Then
            Exit Sub
        End If

        Call despliegaPagosHH
        
        cmdImprimir.Enabled = True
        cmdregistra.Enabled = True
        
    'End If
    
    
'    If MsgBox("Los pagos capturados est�n correctos?", vbQuestion + vbYesNo, "Registro de Pagos") = vbYes Then
'
'        Dim oPago As New Pago
'
'        If oPago.grabaPagos(obtenPagos(False)) Then
'
'            MsgBox "Los pagos fueron grabados", vbInformation + vbOKOnly, "Registro de Pagos"
'            bGrabado = True
'            cmdregistra.Enabled = True
'            cmdgraba.Enabled = False
'            rpagos.PrintReport
'
'        End If
'
'        Set oPago = Nothing
'
'    Else
'        cmdgraba.Enabled = True
'        cmdregistra.Enabled = False
'    End If
    
    'sprPago(MODO_MANUAL).Enabled = True

'    Dim datos As Recordset
'    Dim i As Integer
'    Dim nocliente, folio As Long
'    Dim fecha As Date
'    Dim Pago, adeudo As Double
'    Dim Nombre As String
'
'    base.Execute "delete from pagos_temp"
'
'    For i = 1 To grdpagos.Rows - 2
'
'        grdpagos.Row = i
'        grdpagos.Col = 1
'        folio = CLng(grdpagos.Text)
'        grdpagos.Col = 2
'        nocliente = CLng(grdpagos.Text)
'        grdpagos.Col = 3
'        Nombre = grdpagos.Text
'        grdpagos.Col = 4
'        fecha = CDate(grdpagos.Text)
'        grdpagos.Col = 5
'        Pago = CDbl(grdpagos.Text)
'        grdpagos.Col = 6
'        adeudo = CDbl(grdpagos.Text)
'
'        base.Execute "insert into pagos_temp (no_cliente,factura,fecha,Cantpagada,Cantadeudada,nombre,orden) values(" + CStr(nocliente) + "," + CStr(folio) + ",'" + CStr(fecha) + "'," + CStr(Pago) + "," + CStr(adeudo) + ",'" + Nombre + "'," + CStr(i) + ")"
'
'    Next i
'    MsgBox "Los pagos fueron registrados", vbInformation, "Registro de Pagos"
'    rpagos.PrintReport
'    cmdregistra.Enabled = True
'    cmdgraba.Enabled = False
'    grdpagos.Enabled = False

End Sub

Private Sub cmdregistra_Click()

    If MsgBox("Los pagos est�n correctos?", vbQuestion + vbYesNo, "Registro de Pagos") = vbYes Then

        Dim oPago As New Pago
        
        oPago.registraPagos
        
        MsgBox "Los pagos fueron registrados", vbInformation + vbOKOnly, "Registro de Pagos"
        
        fnLimpiaGrid sprPago(MODO_AUTOMATICO)
        txttotpago.Caption = Format(obtenTotalGrid(sprPago(MODO_AUTOMATICO), COL_PAGO), "$#,####.00")
        'txttotadeudo.Text = Format(obtenTotalGrid(sprPago(MODO_AUTOMATICO), COL_ADEUDO), "###,###,###,###0.00")
        
        cmdregistra.Enabled = False
        cmdImprimir.Enabled = False
        
        'If oPago.grabaPagos(obtenPagos(True)) Then
            
        '    MsgBox "Los pagos fueron registrados", vbInformation + vbOKOnly, "Registro de Pagos"
        '    cmdregistra.Enabled = False
        '    cmdgraba.Enabled = True
            
        '    Call fnLimpiaGrid(sprPago(MODO_MANUAL))
        '    bGrabado = False
            
        'End If
        
        Set oPago = Nothing
        
    Else
        cmdgraba.Enabled = True
        cmdregistra.Enabled = False
    End If
    
    'sprPago(MODO_MANUAL).Enabled = True
            
'    Dim datos As Recordset
'    Dim i As Integer
'    Dim nocliente, folio As Long
'    Dim fecha As Date
'    Dim pago, adeudo As Double
'    Dim hora As String
'
'    If MsgBox("Los pagos registrados est�n correctos?", vbYesNo, "Registro de Pagos") = vbYes Then
'
'        For i = 1 To grdpagos.Rows - 2
'            grdpagos.Row = i
'            grdpagos.Col = 1
'            folio = CLng(grdpagos.Text)
'            grdpagos.Col = 2
'            nocliente = CLng(grdpagos.Text)
'            grdpagos.Col = 4
'            fecha = CDate(grdpagos.Text)
'            grdpagos.Col = 5
'            pago = CDbl(grdpagos.Text)
'            grdpagos.Col = 6
'            adeudo = CDbl(grdpagos.Text)
'            grdpagos.Col = 7
'            hora = CStr(grdpagos.Text)
'            base.Execute "insert into pagos (no_cliente,factura,fecha,Cantpagada,Cantadeudada,orden,cons_pago,usuario,hora, lugar) values(" + CStr(nocliente) + "," + CStr(folio) + ",'" + CStr(fecha) + "'," + CStr(pago) + "," + CStr(adeudo) + "," + CStr(i) + ",1,'" + nombreaux + "','" + CStr(hora) + "','captura')"
'
'        Next i
'
'        ordensigue = i
'
'        MsgBox "Los pagos fueron registrados", vbInformation, "Registro de Pagos"
'        cmdregistra.Enabled = False
'    Else
'        cmdgraba.Enabled = True
'        cmdregistra.Enabled = False
'    End If
'
'    grdpagos.Enabled = True

End Sub

'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub


'Private Sub grdpagos_DblClick()
'
'    Dim totpago1, totadeudo1 As Double
'
'    renglon = grdpagos.Row
'    grdpagos.Col = 1
'    If grdpagos.Text <> "" Then
'        txtfolio.Text = grdpagos.Text
'        grdpagos.Col = 5
'        totpago1 = CDbl(grdpagos.Text)
'        grdpagos.Col = 6
'        totadeudo1 = CDbl(grdpagos.Text)
'        txttotpago.Text = Format(CDbl(txttotpago.Text) - totpago1, "###,###,###,###0.00")
'        txttotadeudo.Text = Format(CDbl(txttotadeudo.Text) - totadeudo1, "###,###,###,###0.00")
'        Call txtfolio_KeyPress(13)
'        cmdagregar.Caption = "Modifica Pago"
'        txtfolio.Enabled = False
'        txtpago.SetFocus
'    End If
'
'End Sub


'Private Function despliegaDatosCredito(cCredito As Collection)
'
'    Dim cRegistro As New Collection
'    Dim oCampo As New Campo
'    Dim fCantidadTotal As Double
'    Dim fCantidadPagada As Double
'
'    Dim fechas, fechaini, pagado As Double
'    Dim pagos As Integer
'    Dim total As Double
'
'    fechas = CDbl(CDate(Format(Now(), "dd/mm/yyyy")))
'
'    Set cRegistro = cCredito(1)
'
'    Set oCampo = cRegistro(10) 'Fecha
'    txtdias.Text = CDate(Format(Now, "dd/mm/yyyy")) - CDate(oCampo.Valor) - 1
'
'    If CDbl(Now) < CDbl(oCampo.Valor) Then
'
'        MsgBox "La fecha de inicio de pago es mayor a hoy", vbInformation, "Registro de Pagos"
'        cmdagregar.Enabled = False
'        txtdias_atraso.Text = ""
'        txtcantatraso.Text = ""
'        txtcte.Text = ""
'        txtnombre.Text = ""
'        txtpago.Text = ""
'        txtadeudo.Text = ""
'        txtfecha.Text = Format(Now, "dd/mm/yyyy")
'        txtfolio.Text = ""
'        txtfolio.SetFocus
'
'        Exit Function
'
'    End If
'
'    Set oCampo = cRegistro(1) 'CLiente
'    txtcte.Text = oCampo.Valor
'
'    Set oCampo = cRegistro(4) 'Cantidad a pagar
'    txtpago.Text = Format(oCampo.Valor, "###,###,###,###0.00")
'    pagogen = Format(oCampo.Valor, "###,###,###,###0.00")
'
'    Set oCampo = cRegistro(6) 'Cantidad total
'    adeudo = Format(oCampo.Valor, "###,###,###,###0.00")
'
'    txtadeudo.Text = Format(adeudo, "###,###,###,###0.00")
'    total = oCampo.Valor
'
'    Set oCampo = cRegistro(8) 'Fecha inicio cr�dito
'    fechaini = CDbl(oCampo)
'
'    Set oCampo = cRegistro(7) 'No de Pagos
'    pagos = oCampo.Valor
'
'    Set oCampo = cRegistro(20) 'Nombre cliente
'    txtnombre.Text = oCampo.Valor
'    Set oCampo = cRegistro(21) 'Apellido cliente
'    txtnombre.Text = txtnombre.Text + " " + oCampo.Valor
'
'    Set oCampo = cRegistro(18) 'Cantidad total
'    fCantidadTotal = oCampo.Valor
'    Set oCampo = cRegistro(19) 'Cantidad total
'    fCantidadPagada = oCampo.Valor
'    txtadeudo.Text = Format(fCantidadTotal - fCantidadPagada, "###,###,###,###0.00")
'    fAdeudo = fCantidadTotal - fCantidadPagada
'
'    txtpago.SelLength = 255
'    txtpago.SetFocus
'
'    txtdias_atraso.Text = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado) / CDbl(txtpago.Text), "###,###,###,###0.00")
'    txtcantatraso.Text = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado), "###,###,###,###0.00")
'
'    If adeudo > 0 Then
'        cmdagregar.Enabled = True
'    Else
'        cmdagregar.Enabled = False
'        MsgBox "La deuda contraida en el folio esta liquidada", vbInformation, "Registro de Pagos"
'    End If
'
'End Function

''Private Sub txtfolio_KeyPress(KeyAscii As Integer)
''
''    Dim oCredito As New credito
''
''    If oCredito.datosCredito(Val(txtfolio.Text)).Count <= 0 Then
''
''        MsgBox "�El folio no existe, verifique por favor!", vbInformation + vbOKOnly
''        txtfolio.SetFocus
''        Exit Sub
''
''    End If
''
''    Call despliegaDatosCredito(oCredito.cDatos)
''
''    Set oCredito = Nothing
    
'    Dim datos As Recordset
'    Dim datos1 As Recordset
'    Dim bandera As Integer
'
'    Dim fechas, fechaini, pagado As Double
'    Dim pagos As Integer
'    Dim total As Double
'
'    fechas = CDbl(CDate(Format(Now(), "dd/mm/yyyy")))
'
'    If KeyAscii = 13 Then
'        If txtfolio.Text <> "" And IsNumeric(txtfolio.Text) Then
'            txtfolio.Text = CLng(txtfolio.Text)
'            Set datos = base.OpenRecordset("select * from creditos where factura=" & CStr(txtfolio.Text))
'            If datos.RecordCount > 0 Then
'                If CDbl(Now) >= CDbl(datos!fecha) Then
'                    txtcte.Text = datos!no_cliente
'                    txtpago.Text = Format(datos!Cantpagar, "###,###,###,###0.00")
'                    txtpago.SelLength = 255
'                    txtpago.SetFocus
'                    pagogen = Format(datos!Cantpagar, "###,###,###,###0.00")
'                    adeudo = Format(datos!Canttotal, "###,###,###,###0.00")
'                    total = datos!Canttotal
'                    txtadeudo.Text = Format(adeudo, "###,###,###,###0.00")
'                    txtdias.Text = CDate(Format(Now, "dd/mm/yyyy")) - CDate(datos!fecha) - 1
'                    If adeudo > 0 Then
'                        cmdagregar.Enabled = True
'                    Else
'                        cmdagregar.Enabled = False
'                        MsgBox "La deuda contraida en el folio esta liquidada", vbInformation, "Registro de Pagos"
'                    End If
'
'                    '***************************************
'                    fechaini = CDbl(datos!fechaini)
'                    pagos = datos!no_pagos
'                    '***************************************
'
'                    datos.Close
'                    Set datos = base.OpenRecordset("select * from clientes where no_cliente=" & CStr(txtcte.Text))
'                    If datos.RecordCount > 0 Then
'                        txtnombre.Text = datos!Nombre + " " + datos!apellido
'                        datos.Close
'                        Set datos = base.OpenRecordset("select * from pagado where no_cliente=" & CStr(txtcte.Text) & " and factura=" & CStr(txtfolio.Text))
'                        If datos.RecordCount > 0 Then
'                            txtadeudo.Text = Format(datos!Canttotal - datos!Cantpagada, "###,###,###,###0.00")
'
'                            adeudo = datos!Canttotal - datos!Cantpagada
'                            pagado = datos!Cantpagada
'                            datos.Close
'                            If adeudo > 0 Then
'                                cmdagregar.Enabled = True
'                            Else
'                                'cmdagregar.Enabled = False
'                                'MsgBox "Deuda liquidada", vbInformation, "Pagos"
'                                'MsgBox "La deuda contraida en el folio esta liquidada", vbInformation, "Registro de Pagos"
'                                'cmdagregar.Enabled = True
'                            End If
'
'                            txtdias_atraso.Text = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado) / CDbl(txtpago.Text), "###,###,###,###0.00")
'                            txtcantatraso.Text = Format((IIf((fechas - fechaini) > pagos, pagos, (fechas - fechaini)) * (total / pagos) - pagado), "###,###,###,###0.00")
'
'                        Else
'                            datos.Close
'
'                        End If
'                    End If
'                Else
'                    MsgBox "La fecha de inicio de pago es mayor a hoy", vbInformation, "Registro de Pagos"
'                    cmdagregar.Enabled = False
'                    txtdias_atraso.Text = ""
'                    txtcantatraso.Text = ""
'                    txtcte.Text = ""
'                    txtnombre.Text = ""
'                    txtpago.Text = ""
'                    txtadeudo.Text = ""
'                    txtfecha.Text = Format(Now, "dd/mm/yyyy")
'                    txtfolio.Text = ""
'                    txtfolio.SetFocus
'                End If
'
'            Else
'                MsgBox "El folio que se esta buscando no existe", vbInformation, "Registro de Pagos"
'                cmdagregar.Enabled = False
'                txtdias_atraso.Text = ""
'                txtcantatraso.Text = ""
'                txtcte.Text = ""
'                txtnombre.Text = ""
'                txtpago.Text = ""
'                txtadeudo.Text = ""
'                txtfecha.Text = Format(Now, "dd/mm/yyyy")
'                txtfolio.Text = ""
'                txtfolio.SetFocus
'
'            End If
'        Else
'        End If
'
'    End If
    
''End Sub

'Private Sub txtfolio_LostFocus()
'    txtfolio_KeyPress (13)
'End Sub

'Private Sub txtpago_GotFocus()
'
'    If txtfolio.Text = "" Or Not (IsNumeric(txtfolio.Text)) Then
'        txtfolio.Text = ""
'        txtfolio.SetFocus
'    End If
'
'End Sub

'Private Sub txtpago_LostFocus()
'
'    If txtfolio.Text <> "" And IsNumeric(txtfolio.Text) Then
'        If txtpago.Text <> "" Then
'            If txtpago.Text <> "" And IsNumeric(txtpago.Text) Then
'                If adeudo > CDbl(txtpago.Text) Then
'                    txtadeudo.Text = Format(adeudo - CDbl(txtpago.Text), "###,###,###,###0.00")
'                Else
'                    txtadeudo.Text = Format(adeudo - CDbl(txtpago.Text), "###,###,###,###0.00")
'                End If
'            Else
'                MsgBox "El dato del pago es incorrecto", vbCritical, "Registro de Pagos"
'                txtpago.Text = Format(pagogen, "###,###,###,###0.00")
'                txtpago.SelLength = 255
'                txtpago.SetFocus
'            End If
'            bandera = 1
'        Else
'            MsgBox "El dato del pago es incorrecto", vbCritical, "Registro de Pagos"
'            txtpago.Text = Format(pagogen, "###,###,###,###0.00")
'            txtpago.SetFocus
'        End If
'        bandera = 1
'    End If
'
'End Sub

'Private Sub cmdagregar_Click()
'
'Dim i As Integer
'Dim folio, Cliente As Long
'Dim totadeudo, totpago As Double
'
'If bandera = 0 Then
'    Call txtpago_LostFocus
'End If
'
'totpago = 0
'totadeudo = 0
'
'If cmdagregar.Caption = "Modifica Pago" Then
'    grdpagos.Row = renglon
'    If txtpago.Text <> "" And IsNumeric(txtpago.Text) Then
'        grdpagos.Col = 1
'        grdpagos.Text = txtfolio.Text
'        grdpagos.Col = 2
'        grdpagos.Text = txtcte.Text
'        grdpagos.Col = 3
'        grdpagos.Text = txtnombre.Text
'        grdpagos.Col = 4
'        grdpagos.Text = txtfecha.Text
'        grdpagos.Col = 5
'        grdpagos.Text = Format(txtpago.Text, "###,###,###,###0.00")
'        grdpagos.Col = 6
'        grdpagos.Text = txtadeudo.Text
'        grdpagos.Col = 7
'        grdpagos.Text = Format(Now, "hh:mm AM/PM")
'        txtfolio.Enabled = True
'        cmdgraba.Enabled = True
'    Else
'        MsgBox "El dato del pago es incorrecto", vbCritical, "Registro de Pagos"
'        txtpago.Text = pagogen
'        txtpago.SetFocus
'    End If
'End If
'
'If grdpagos.Rows > 2 Then
'    For i = 1 To grdpagos.Rows - 2
'    grdpagos.Row = i
'    grdpagos.Col = 1
'    folio = CLng(grdpagos.Text)
'    grdpagos.Col = 2
'    Cliente = CLng(grdpagos.Text)
'    If Cliente = CLng(txtcte.Text) And folio = CLng(txtfolio.Text) And cmdagregar.Caption <> "Modifica Pago" Then
'        MsgBox "El pago ya fue registrado", vbInformation, "Registro de pagos"
'        GoTo fin
'    Else
'        grdpagos.Col = 5
'        totpago = totpago + CDbl(grdpagos.Text)
'        grdpagos.Col = 6
'        totadeudo = totadeudo + CDbl(grdpagos.Text)
'        grdpagos.Col = 7
'    End If
'    Next i
'End If
'Dim fecha As Date
'Dim datos2 As Recordset
'
'fecha = Format(Now, "dd/mm/yyyy")
'
'
'Set datos2 = base.OpenRecordset("select factura from pagos where factura= " + CStr(txtfolio.Text) + " and CStr(fecha) ='" + CStr(fecha) + "'")
'
'If datos2.RecordCount > 0 Then
'    MsgBox "Ya existe un pago con la fecha de hoy para el folio " + txtfolio.Text + "", vbOKOnly, "Pagos"
'    Exit Sub
'End If
'
'
'If cmdagregar.Caption <> "Modifica Pago" Then
'    If txtfolio.Text <> "" And IsNumeric(txtfolio.Text) Then
'        If txtpago.Text <> "" And IsNumeric(txtpago.Text) Then
'            grdpagos.Row = grdpagos.Rows - 1
'            grdpagos.Rows = grdpagos.Row + 2
'            grdpagos.Col = 1
'            grdpagos.Text = txtfolio.Text
'            grdpagos.Col = 2
'            grdpagos.Text = txtcte.Text
'            grdpagos.Col = 3
'            grdpagos.Text = txtnombre.Text
'            grdpagos.Col = 4
'            grdpagos.Text = txtfecha.Text
'            grdpagos.Col = 5
'            grdpagos.Text = Format(txtpago.Text, "###,###,###,###0.00")
'            grdpagos.Col = 6
'            grdpagos.Text = txtadeudo.Text
'            grdpagos.Col = 5
'            totpago = totpago + CDbl(grdpagos.Text)
'            grdpagos.Col = 6
'            totadeudo = totadeudo + CDbl(grdpagos.Text)
'            txttotpago.Text = Format(totpago, "###,###,###,###0.00")
'            txttotadeudo.Text = Format(totadeudo, "###,###,###,###0.00")
'            cmdgraba.Enabled = True
'            txtultimo.Text = txtfolio.Text
'            txtpago1.Text = Format(txtpago.Text, "###,###,###,###0.00")
'            txtadeudo1.Text = txtadeudo.Text
'            grdpagos.Col = 7
'            grdpagos.Text = Format(Now, "hh:mm AM/PM")
'        Else
'            MsgBox "El dato del pago es incorrecto", vbCritical, "Registro de Pagos"
'            txtpago.Text = pagogen
'            txtpago.SetFocus
'        End If
'    Else
'        MsgBox "El dato del folio es incorrecto", vbCritical, "Registro de Pagos"
'        txtfolio.SetFocus
'    End If
'Else
'    txttotpago.Text = Format(totpago, "###,###,###,###0.00")
'    txttotadeudo.Text = Format(totadeudo, "###,###,###,###0.00")
'End If
'
'fin:
'bandera = 0
'cmdagregar.Caption = "Agregar"
'cmdagregar.Enabled = False
'txtdias_atraso.Text = ""
'txtcantatraso.Text = ""
'txtcte.Text = ""
'txtnombre.Text = ""
'txtpago.Text = ""
'txtadeudo.Text = ""
'txtfecha.Text = Format(Now, "dd/mm/yyyy")
'txtfolio.Text = ""
'txtfolio.SetFocus
'End Sub

