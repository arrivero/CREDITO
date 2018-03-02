VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form sobrantesfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobrantes"
   ClientHeight    =   3105
   ClientLeft      =   4455
   ClientTop       =   3270
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optsobrante 
      Caption         =   "Sobrante"
      Height          =   255
      Left            =   750
      TabIndex        =   0
      Top             =   1200
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optfaltante 
      Caption         =   "Faltante"
      Height          =   255
      Left            =   2190
      TabIndex        =   1
      Top             =   1200
      Width           =   885
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   690
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtfecha 
      Height          =   345
      Left            =   1350
      TabIndex        =   6
      Top             =   360
      Width           =   1755
      _Version        =   65537
      _ExtentX        =   3096
      _ExtentY        =   609
      _StockProps     =   93
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin EditLib.fpCurrency txtcantidad 
      Height          =   345
      Left            =   1500
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   609
      Enabled         =   -1  'True
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
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   390
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   750
      TabIndex        =   4
      Top             =   1695
      Width           =   735
   End
End
Attribute VB_Name = "sobrantesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nombreglobal As String
Dim bgral As Integer

Private Sub Form_Load()

    txtfecha.Text = Format(Now, "dd/mm/yyyy")
    
    txtcantidad.Text = 0
    
    Me.Caption = "Sobrantes " + UCase(gstrUsuario)

End Sub

Private Sub cmdagregar_Click()
    
    Dim tipo As String
    Dim cantidad As Double
    
    cantidad = Val(fnstrValor(txtcantidad.Text))
    
    If cantidad <= 0# Then
        MsgBox "¡La cantidad no es correca, verifique por favor!", vbInformation + vbOKOnly
        txtcantidad.Text = 0#
        
        'txtcantidad.SetFocus
        Exit Sub
    End If
    
    If optsobrante.Value = True Then
        tipo = "S"
    Else
        If optfaltante.Value = True Then
            tipo = "F"
        End If
    End If
    
    Dim oPago As New Pago
    oPago.registraSobrante txtfecha.Text, tipo, cantidad, gstrUsuario
    Set oPago = Nothing
    
    'Base.Execute "insert into sobrantes (fecha,tipo,cantidad,usuario) values('" + Format(Now(), "dd/mm/yyyy") + "','" + tipo + "'," + CStr(cantidad) + ",'" + nombreaux + "')"
    
    MsgBox "El dato fue agregado", vbOKOnly, "Sobrantes"
        
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

