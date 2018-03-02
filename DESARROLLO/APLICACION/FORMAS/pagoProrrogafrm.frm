VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form pagoProrrogafrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo Pendiente"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegistro 
      Caption         =   "Registro Pago"
      Height          =   465
      Left            =   2130
      TabIndex        =   6
      Top             =   1350
      Width           =   1215
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtFechaPago 
      Height          =   345
      Left            =   1170
      TabIndex        =   2
      Top             =   780
      Width           =   1725
      _Version        =   65537
      _ExtentX        =   3043
      _ExtentY        =   609
      _StockProps     =   93
      MinDate         =   "2007/1/1"
      MaxDate         =   "2010/12/31"
      Mask            =   2
   End
   Begin EditLib.fpDoubleSingle fTotalPorPagar 
      Height          =   345
      Left            =   4020
      TabIndex        =   4
      Top             =   780
      Width           =   1725
      _Version        =   196608
      _ExtentX        =   3043
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegFormat       =   1
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
   Begin VB.Label Label6 
      Caption         =   "Total a pagar:"
      Height          =   255
      Left            =   2940
      TabIndex        =   5
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha de pago:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   810
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto:"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   2295
   End
   Begin VB.Label lblConcepto 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   5835
   End
End
Attribute VB_Name = "pagoProrrogafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iPago As Integer

Private Sub Form_Load()

    dtFechaPago.MinDate = DateAdd("d", 1, Date)
    
End Sub

Private Sub cmdRegistro_Click()
    
    Dim dMonto As Double
    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    Set cCheque = New Collection
    
    dMonto = Val(fnstrValor(fTotalPorPagar.Text))
    
    'cCheque.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
    cCheque.Add oCampo.CreaCampo(adInteger, , , dtFechaPago.Text)
    cCheque.Add oCampo.CreaCampo(adInteger, , , dMonto)
    cCheque.Add oCampo.CreaCampo(adInteger, , , iPago)
    
    cCheques.Add cCheque
        
    Dim oProveedor As New cProveedor
    
    Call oProveedor.pagoRegistraProrroga(cCheques)
    Set oProveedor = Nothing
    Unload Me
    
End Sub

