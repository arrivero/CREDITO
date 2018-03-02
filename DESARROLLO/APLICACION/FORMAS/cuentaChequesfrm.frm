VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form cuentaChequesfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta de Cheques"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame2 
      Height          =   2355
      Left            =   4740
      TabIndex        =   11
      Top             =   60
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4154
      _Version        =   196608
      ForeColor       =   -2147483635
      Caption         =   "Origen de los recursos"
      Begin VB.ComboBox cbOtrasFuentas 
         Height          =   315
         Left            =   150
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.ComboBox cbCuentaCheques 
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Top             =   510
         Width           =   3105
      End
      Begin VB.Label lblOtrasFuentes 
         Caption         =   "Otras fuentes:"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   930
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Cuenta de Cheques"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   7020
      TabIndex        =   1
      Top             =   2490
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   465
      Left            =   5730
      TabIndex        =   0
      Top             =   2490
      Width           =   1215
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2355
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4154
      _Version        =   196608
      ForeColor       =   -2147483635
      Caption         =   "Datos"
      Begin VB.TextBox txtSaldoCtaCheqes 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   2
         EndProperty
         Height          =   330
         Left            =   3105
         TabIndex        =   18
         Top             =   1170
         Width           =   1500
      End
      Begin VB.ComboBox cbBanco 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   510
         Width           =   2325
      End
      Begin VB.TextBox txtCuenta 
         Height          =   345
         Left            =   150
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1200
         Width           =   2325
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   345
         Left            =   150
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1920
         Width           =   4455
      End
      Begin EditLib.fpDoubleSingle txtSaldoCtaCheqesold 
         Height          =   345
         Left            =   3105
         TabIndex        =   6
         Top             =   1170
         Width           =   1515
         _Version        =   196608
         _ExtentX        =   2672
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
         ControlType     =   0
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
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Número de la Cuenta"
         Height          =   285
         Left            =   150
         TabIndex        =   9
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción de la Cuenta"
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label Label8 
         Caption         =   "Saldo inicial:"
         Height          =   195
         Left            =   3090
         TabIndex        =   7
         Top             =   930
         Width           =   1425
      End
   End
   Begin VB.Label Label5 
      Caption         =   "2. Definir la fuente de ingresos, del saldo inicial de la cuenta de cheques nueva."
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   2790
      Width           =   5685
   End
   Begin VB.Label Label4 
      Caption         =   "1. Definir los datos generales de la cuenta de cheques nueva."
      Height          =   285
      Left            =   60
      TabIndex        =   12
      Top             =   2490
      Width           =   5505
   End
End
Attribute VB_Name = "cuentaChequesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim oBanco As New Banco
    If oBanco.fnCatalogo Then
        Call fnLlenaComboCollecion(cbBanco, oBanco.cDatos, 0, "")
    End If
    Set oBanco = Nothing

    Dim oCtaCheques As New CuentaCheques
    If oCtaCheques.catalogoEsp = True Then
        Call fnLlenaComboCollecion(cbCuentaCheques, oCtaCheques.cDatos, 0, "")
    End If
    Set oCtaCheques = Nothing
    
'    If oAlmacen.contabilidadAbilitada = "SI" Then
'        lblOtrasFuentes.Visible = True
'        cbOtrasFuentas.Visible = True
'        'Carga catálogo FUENTES DE INGRESO
'        Dim oProveedor As New cProveedor
'        If oProveedor.catalogoGastos(ID_OPERACION_CAPITAL_SOCIAL) Then
'            Call fnLlenaComboCollecion(cbOtrasFuentas, oProveedor.cDatos, 0)
'        End If
'        'Con el concepto y el tipo de operación se obtiene la cuenta.
'        Set oProveedor = Nothing
'    Else
'        lblOtrasFuentes.Visible = False
'        cbOtrasFuentas.Visible = False
'    End If
    
End Sub

Private Sub cmdAceptar_Click()
    
    'valida si no es un nombre de banco nuevo
    Dim iBanco As Integer
    Dim strBanco As String
    'Dim strMensaje As String
    
    iBanco = esBancoNuevofn(strBanco)
        
    If iBanco > 0 Then
        Dim cCuentaBanco As New Collection
        Dim Registro As New Collection
        Dim oCampo As New Campo
    
        'Registro.Add oCampo.CreaCampo(adInteger, , , gAlmacen)
        Registro.Add oCampo.CreaCampo(adInteger, , , iBanco)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtCuenta)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtDescripcion)
        Registro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtSaldoCtaCheqes)))

        cCuentaBanco.Add Registro
        
        Dim oCuentaCheques As New CuentaCheques
        
        'strMensaje = oCuentaCheques.creaCuenta(gAlmacen, cCuentaBanco)
        Call oCuentaCheques.creaCuenta(cCuentaBanco)
        'If strMensaje = "YA_HAY_MODULO_CONTABLE" Then
        '    MsgBox "La cuenta de cheques se registró con éxito" + Chr(13) + "Es muy importante que relacione esta con la cuenta contable que afectará, antes de hacer cualquier movimiento con la cuenta de cheques!", vbInformation + vbOKOnly
        'Else
            MsgBox "La cuenta de cheques se registró con éxito", vbInformation
        'End If
        Unload Me
    Else
        MsgBox "No fue posible crear la cuenta de cheques", vbCritical + vbInformation
    End If
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Function esBancoNuevofn(ByRef strBanco As String) As Integer
    
    If cbBanco.ListIndex = -1 Then  ' banco nuevo
        
        'busca en base de datos si no existe ya otro con el mismo nombre
        Dim oBanco As New Banco
        
        If oBanco.buscaConNombre(cbBanco.Text) = False Then
            
            'CREA EL BANCO Y OBTEN EL IDENTIFICADOR
            Dim iBanco As Integer
            iBanco = oBanco.creaNuevo(cbBanco.Text)
                   
        End If
        
    Else
    
        iBanco = cbBanco.ItemData(cbBanco.ListIndex)
        
    End If
    
    Set oBanco = Nothing
    
    esBancoNuevofn = iBanco
    
End Function

