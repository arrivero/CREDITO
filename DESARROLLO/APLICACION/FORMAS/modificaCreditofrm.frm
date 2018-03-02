VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form modificaCreditofrm 
   Caption         =   "Modificación de Créditos"
   ClientHeight    =   7335
   ClientLeft      =   1935
   ClientTop       =   2655
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox txtadeudo 
      Height          =   375
      Left            =   5940
      TabIndex        =   43
      Top             =   6030
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   255
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ ######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtpagado 
      Height          =   375
      Left            =   5940
      TabIndex        =   42
      Top             =   5625
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   32768
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ ######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txttotalpagar 
      Height          =   375
      Left            =   2025
      TabIndex        =   41
      Top             =   6030
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ ########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtnotadescuento 
      Height          =   330
      Left            =   2055
      TabIndex        =   40
      Top             =   5625
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ #####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtdescuento 
      Height          =   330
      Left            =   2055
      TabIndex        =   39
      Top             =   5175
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ #####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtnotacred 
      Height          =   330
      Left            =   2055
      TabIndex        =   38
      Top             =   4770
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ #####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtMora 
      Height          =   330
      Left            =   2055
      TabIndex        =   37
      Top             =   4320
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ #####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtpagodiario 
      Height          =   330
      Left            =   2055
      TabIndex        =   36
      Top             =   3870
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "$ #####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtpagos 
      Height          =   330
      Left            =   2055
      TabIndex        =   35
      Top             =   3420
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtfinan 
      Height          =   330
      Left            =   2055
      TabIndex        =   34
      Top             =   2970
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
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
   Begin MSMask.MaskEdBox txtfinanpor 
      Height          =   330
      Left            =   2055
      TabIndex        =   33
      Top             =   2520
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtcredito 
      Height          =   330
      Left            =   2055
      TabIndex        =   32
      Top             =   2115
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   393216
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
   Begin MSMask.MaskEdBox txtfactura 
      Height          =   420
      Left            =   2055
      TabIndex        =   31
      Top             =   1350
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   -2147483635
      PromptInclude   =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtnocliente 
      Height          =   375
      Left            =   1350
      TabIndex        =   30
      Top             =   180
      Width           =   1320
      _ExtentX        =   2328
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
   Begin VB.TextBox txtnombre 
      Height          =   330
      Left            =   1350
      MaxLength       =   255
      TabIndex        =   29
      Top             =   630
      Width           =   4200
   End
   Begin VB.TextBox txtdescrip 
      Height          =   765
      Left            =   2610
      TabIndex        =   20
      Top             =   6615
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox opcion 
      Caption         =   "Electricos"
      Height          =   255
      Left            =   270
      TabIndex        =   19
      Top             =   6705
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbstatus 
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
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3870
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6630
      TabIndex        =   2
      Top             =   6735
      Width           =   1335
   End
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "Actualiza"
      Height          =   495
      Left            =   5190
      TabIndex        =   1
      Top             =   6735
      Width           =   1335
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtfecha 
      Height          =   345
      Left            =   6480
      TabIndex        =   27
      Top             =   2115
      Width           =   1455
      _Version        =   65537
      _ExtentX        =   2566
      _ExtentY        =   609
      _StockProps     =   93
      Enabled         =   0   'False
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtfechaini 
      Height          =   345
      Left            =   6480
      TabIndex        =   28
      Top             =   2550
      Width           =   1455
      _Version        =   65537
      _ExtentX        =   2566
      _ExtentY        =   609
      _StockProps     =   93
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin VB.Label txtdias 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6435
      TabIndex        =   45
      Top             =   4815
      Width           =   1455
   End
   Begin VB.Label txtfechafin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6480
      TabIndex        =   44
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   90
      X2              =   7950
      Y1              =   6570
      Y2              =   6570
   End
   Begin VB.Label Label22 
      Caption         =   "Nota de Descuento:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label21 
      Caption         =   "Descuento:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label20 
      Caption         =   "Nota de Credito:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4830
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "Moratorios:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4380
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "Días de Crédito:"
      Height          =   255
      Left            =   4260
      TabIndex        =   22
      Top             =   4830
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Descripción:"
      Height          =   255
      Left            =   1485
      TabIndex        =   21
      Top             =   6750
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Cantidad Adeudada:"
      Height          =   255
      Left            =   4260
      TabIndex        =   18
      Top             =   6090
      Width           =   1560
   End
   Begin VB.Label Label13 
      Caption         =   "Cantidad Pagada:"
      Height          =   255
      Left            =   4260
      TabIndex        =   17
      Top             =   5640
      Width           =   1560
   End
   Begin VB.Label Label16 
      Caption         =   "Financiamiento en %:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2625
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "%"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   2610
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "No. de Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Folio:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Monto del crédito:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2130
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Financiamiento:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3075
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "No. de pagos:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad a pagar:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3930
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Cantidad total a pagar:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6090
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha inicio de cobranza:"
      Height          =   255
      Left            =   4260
      TabIndex        =   6
      Top             =   2610
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Fecha de cambio de status:"
      Height          =   255
      Left            =   4260
      TabIndex        =   5
      Top             =   3090
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Fecha de registro:"
      Height          =   255
      Left            =   4260
      TabIndex        =   4
      Top             =   2175
      Width           =   2010
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label12 
      Caption         =   "Status:"
      Height          =   255
      Left            =   4260
      TabIndex        =   3
      Top             =   3885
      Width           =   615
   End
End
Attribute VB_Name = "modificaCreditofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim anterior As Double
Dim estadocre As Integer

Public iNoFactura As Long
Public iNoCliente As Integer
Public strNombreCliente As String
Public bModifica As Boolean

Private Sub Form_Load()

    'Carga catálogo estados de crédito
    Dim oCredito As New credito
    If oCredito.statusCatalogo() Then
        Call fnLlenaComboCollecion(cmbstatus, oCredito.cDatos, 0, "")
    End If
    
    'Buscar los datos del credito, enviando el No. de cliente y el No. de Factura
    'Desplegar los datos del crédito (el desplegado se debe hacer en la función despliegaDatos)
    Call despliegaDatos(oCredito.obtenGenerales(iNoFactura))
    
    If giTipoUsuario <> USUARIO_GERENTE Then
        cmdgrabar.Enabled = False
        habilitaControles False
    Else
        habilitaControles bModifica
    End If
    
    Set oCredito = Nothing
    
End Sub

Private Function habilitaControles(bModifica As Boolean)

        txtnocliente.Enabled = bModifica
        
        'txtfactura.Enabled = bModifica
        
        txtnombre.Enabled = bModifica
        
        txtnombre.Enabled = bModifica
        
        txtcredito.Enabled = bModifica
        txtfinan.Enabled = bModifica
        
        txtpagos.Enabled = bModifica
        
        txtpagodiario.Enabled = False 'bModifica
        
        txttotalpagar.Enabled = False 'bModifica
        
        'txtfecha.Enabled = bModifica
        
        txtfechaini.Enabled = bModifica
        
        txtfechafin.Enabled = bModifica
        
        txtMora.Enabled = bModifica
        
        txtnotacred.Enabled = bModifica
        
        txtdescuento.Enabled = bModifica
        
        txtnotadescuento.Enabled = bModifica
        
        txtdescrip.Enabled = bModifica
        
        txtfinanpor.Enabled = bModifica
        
        cmbstatus.Enabled = bModifica
        'txtdias.Enabled = False 'bModifica
            
        txtpagado.Enabled = False 'bModifica
        txtadeudo.Enabled = False 'bModifica

End Function

Private Sub txtfactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        Dim oCredito As New credito
        Dim cDatos As New Collection
        
        Set cDatos = oCredito.obtenGenerales(Val(txtfactura.Text))
        
        If cDatos.Count > 0 Then
            Call despliegaDatos(cDatos)
        Else
            MsgBox "No existe el Crédito " & txtfactura.Text & " verifique por favor!", vbInformation + vbOKOnly
        End If
        
        Set oCredito = Nothing
        
    End If
    
End Sub

Private Function despliegaDatos(cCredito As Collection)

    If IsNull(cCredito) Then
        Exit Function
    End If
    
    If cCredito.Count > 0 Then
    
    
        Dim cRegistro As New Collection
        Dim oCampo As New Campo
        
        Dim fCredito As Double
        Dim fFinanciamiento As Double
        Dim fMoratorios As Double
        Dim fNotaDescuento As Double
        Dim fNotaCredito As Double
        Dim fDescuento As Double
        
        Set cRegistro = cCredito(1)
    
        Set oCampo = cRegistro(1)
        txtnocliente.Text = oCampo.Valor
        
        Set oCampo = cRegistro(2)
        txtfactura.Text = oCampo.Valor
        
        Set oCampo = cRegistro(3)
        txtnombre.Text = oCampo.Valor
        
        Set oCampo = cRegistro(4)
        txtnombre.Text = txtnombre.Text + " " + oCampo.Valor
        
        Set oCampo = cRegistro(5)
        txtcredito.Text = oCampo.Valor ' / 100# 'Format(oCampo.Valor, "###,###,###,###0.00")
        fCredito = Val(fnstrValor(txtcredito.Text))
        anterior = oCampo.Valor
        
        Set oCampo = cRegistro(6)
        txtfinan.Text = oCampo.Valor 'Format(oCampo.Valor, "###,###,###,###0.00")
        fFinanciamiento = Val(fnstrValor(txtfinan.Text))
        
        Set oCampo = cRegistro(7)
        txtpagos.Text = oCampo.Valor
        
        Set oCampo = cRegistro(8)
        txtpagodiario.Text = oCampo.Valor 'Format(oCampo.Valor, "###,###,###,###0.00")
        
        Set oCampo = cRegistro(10)
        txtfecha.Text = oCampo.Valor
        
        Set oCampo = cRegistro(11)
        If IsNull(oCampo.Valor) Then
            txtfechaini.Text = Date
        Else
            txtfechaini.Text = oCampo.Valor
        End If
        
        Set oCampo = cRegistro(12)
        txtfechafin.Caption = oCampo.Valor
        
        Set oCampo = cRegistro(13)
        If IsNull(oCampo.Valor) Then
            fMoratorios = 0#
        Else
            fMoratorios = oCampo.Valor
        End If
        txtMora.Text = fMoratorios 'Format(oCampo.Valor, "###,###,###,###0.00")
                                
        Set oCampo = cRegistro(14)
        If IsNull(oCampo.Valor) Then
            fNotaCredito = 0#
        Else
        fNotaCredito = oCampo.Valor
        End If
        txtnotacred.Text = fNotaCredito 'Format(oCampo.Valor, "###,###,###,###0.00")
        
        Set oCampo = cRegistro(15)
        If IsNull(oCampo.Valor) Then
            fDescuento = 0#
        Else
            fDescuento = oCampo.Valor
        End If
        txtdescuento.Text = fDescuento 'Format(oCampo.Valor, "###,###,###,###0.00")
        
        Set oCampo = cRegistro(16)
        If IsNull(oCampo.Valor) Then
            fNotaDescuento = 0#
        Else
            fNotaDescuento = oCampo.Valor
        End If
        txtnotadescuento.Text = fNotaDescuento 'Format(oCampo.Valor, "###,###,###,###0.00")
        
        'Set oCampo = cRegistro(9)
        txttotalpagar.Text = fCredito + fFinanciamiento + fMoratorios - fNotaCredito - fDescuento - fNotaDescuento 'Format(fCredito + fFinanciamiento + fMoratorios - fNotaCredito - fDescuento - fNotaDescuento, "###,###,###,###0.00")
        
        Set oCampo = cRegistro(24)
        opcion.Value = IIf(IsNull(oCampo.Valor), 0, oCampo.Valor)
        If opcion.Value = 1 Then
            txtdescrip.Visible = True
            'txtdescrip.Text = datos!descripcion
        Else
            txtdescrip.Visible = False
            'txtdescrip.Text = ""
        End If
        
        Set oCampo = cRegistro(20)
        txtfinanpor.Text = oCampo.Valor
        
        Set oCampo = cRegistro(21)
        cmbstatus.ListIndex = fnBuscaIndiceCombo(cmbstatus, oCampo.Valor)
        estadocre = oCampo.Valor
        
        Set oCampo = cRegistro(22)
        txtdias.Caption = oCampo.Valor
            
        Set oCampo = cRegistro(19)
        txtpagado.Text = oCampo.Valor 'Format(oCampo.Valor, "###,###,###,###0.00")
        
        Set oCampo = cRegistro(23)
        txtadeudo.Text = oCampo.Valor 'Format(Val(fnstrValor(txttotalpagar.Text)) - Val(fnstrValor(txtpagado.Text)), "###,###,###,###0.00")
            
    End If
    
End Function

Private Sub cmbstatus_Click()

    If estadocre = 0 And cmbstatus.ListIndex = 2 Then
        txtfechafin.Caption = Date
    End If

End Sub

Private Function ObtenDatosCredito() As Collection

    Dim cCredito As New Collection
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    Dim dTotalCredito As Double
    
    dTotalCredito = Val(fnstrValor(txtcredito.Text)) + Val(fnstrValor(txtfinan.Text))
    
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtnocliente.Text))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtfactura.Text))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtcredito.Text)))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtnotadescuento.Text)))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtnotacred.Text)))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtpagodiario.Text)))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtfinan.Text)))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , dTotalCredito)
    'cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txttotalpagar.Text)))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(txtpagos.Text))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , txtfechaini.Text)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , txtfechafin.Caption)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , txtfecha.Text)

    Select Case cmbstatus.ListIndex
        Case Is = 0
            cRegistro.Add oCampo.CreaCampo(adInteger, , , "V")
        Case Is = 1
            cRegistro.Add oCampo.CreaCampo(adInteger, , , "T")
        Case Is = 2
            cRegistro.Add oCampo.CreaCampo(adInteger, , , "P")
        Case Is = 3
            cRegistro.Add oCampo.CreaCampo(adInteger, , , "C")
        Case Is = 4
            cRegistro.Add oCampo.CreaCampo(adInteger, , , "E")
    End Select
    
    If opcion.Value = 1 Then
        cRegistro.Add oCampo.CreaCampo(adInteger, , , 1)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , txtdescrip.Text)
    Else
        cRegistro.Add oCampo.CreaCampo(adInteger, , , 0)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , "")
    End If
    cRegistro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(txtnotadescuento.Text)))
    
    cRegistro.Add oCampo.CreaCampo(adInteger, , , convierteMontoConLentra(CStr(Val(fnstrValor(txtcredito.Text)))))
    cRegistro.Add oCampo.CreaCampo(adInteger, , , convierteMontoConLentra(CStr(dTotalCredito)))

    cCredito.Add cRegistro
    
    Set ObtenDatosCredito = cCredito
    
End Function

Private Sub cmdgrabar_Click()

        Dim oCredito As New credito
        
        If validaForma = False Then
            Exit Sub
        End If
        
        If oCredito.validaCreditoCliente(Val(txtnocliente.Text), Val(fnstrValor(txtcredito.Text)), anterior) Then
            
            oCredito.actualizaCredito ObtenDatosCredito
            
        Else
        
            MsgBox "¡La cantidad solicitada excede el límite de crédito otorgado al cliente.!", vbInformation + vbOKOnly, "Crédito de Clientes"
            txtcredito.SetFocus
            
        End If
        
        Set oCredito = Nothing
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub opcion_Click()
    
    txtdescrip.Text = ""
    If opcion.Value = 1 Then
        txtdescrip.Visible = True
    End If
    If opcion.Value = 0 Then
        txtdescrip.Visible = False
    End If
    
End Sub

Private Sub txtcredito_GotFocus()
    txtcredito.SelLength = 255
End Sub

Private Function validaForma() As Boolean

    validaForma = False
    
    If Val(fnstrValor(txtpagodiario.Text)) <= 0 Then
        MsgBox "¡El monto de pago diario debe ser mayor a cero!", vbInformation + vbOKOnly, "Crédito del cliente"
        txtcredito.SetFocus
        Exit Function
    End If
    
    If txtfinanpor.Text = "" Then
        MsgBox "El porcentaje a finaciar debe ser mayor a cero", vbInformation, "Crédito del cliente"
        txtfinanpor.SetFocus
        Exit Function
    End If
    
    If Val(fnstrValor(txttotalpagar.Text)) <= 0 Then
        MsgBox "¡El monto total a pagar debe ser mayor a cero!", vbInformation + vbOKOnly, "Crédito del cliente"
        txtcredito.SetFocus
        Exit Function
    End If
        
    If opcion.Value = 1 Then
        If txtdescrip.Text <> "" Then
        Else
            MsgBox "¡Falta capturar la descripción del articulo de venta!", vbCritical + vbOKOnly, "Folios"
            txtdescrip.SetFocus
        End If
    End If
        
    validaForma = True
    
End Function

Private Function pagoDiario() As Double

    Dim fCredito As Double
    Dim fFinanciamiento As Double
    Dim fPagoDiario As Double
    
    Dim strDecimales, strPagoDiario, strEntero As String
    Dim iPosPunto As Integer

    fCredito = Val(fnstrValor(txtcredito.Text))
    fFinanciamiento = Val(fnstrValor(txtfinanpor.Text))
    
    fPagoDiario = (fCredito + (fCredito * (fFinanciamiento / 100))) / Val(txtpagos.Text)
    
    strPagoDiario = CStr(fPagoDiario)
        
    iPosPunto = InStr(strPagoDiario, ".")
    
    If iPosPunto > 0 Then
        strDecimales = Mid(strPagoDiario, iPosPunto + 1, 2)
        
        If Val(strDecimales) > 0 Then
            strEntero = Mid(strPagoDiario, 1, iPosPunto - 1)
            fPagoDiario = Val(strEntero) + 1
        End If
    End If

    pagoDiario = fPagoDiario
    
End Function

Private Function totalAPagar() As Double

    Dim fCredito As Double
    Dim fFinanciamiento As Double
    Dim fMoratorios As Double
    Dim fNotaCredito As Double
    Dim fNotaDescuento As Double
    Dim fDescuento As Double
    
    fCredito = Val(fnstrValor(txtcredito.Text))
    fFinanciamiento = Val(fnstrValor(txtfinanpor.Text))
    fMoratorios = Val(fnstrValor(txtMora.Text))
    fNotaCredito = Val(fnstrValor(txtnotacred.Text))
    fNotaDescuento = Val(fnstrValor(txtnotadescuento.Text))
    fDescuento = Val(fnstrValor(txtdescuento.Text))
    
    totalAPagar = fCredito + (fCredito * (fFinanciamiento / 100)) + fMoratorios - fNotaCredito - fNotaDescuento - fDescuento
    
End Function

Private Function financiamiento() As Double

    Dim fCredito As Double
    Dim fProcentajeFinanciamiento As Double
    
    fCredito = Val(fnstrValor(txtcredito.Text))
    fProcentajeFinanciamiento = Val(fnstrValor(txtfinanpor.Text))
    
    financiamiento = fCredito * (fProcentajeFinanciamiento / 100)
    
End Function

Private Function porcentajeFinanciamiento() As Integer

    Dim fCredito As Double
    Dim fFinanciamiento As Double
    
    fCredito = Val(fnstrValor(txtcredito.Text))
    fFinanciamiento = Val(fnstrValor(txtfinan.Text))
    
    porcentajeFinanciamiento = (fFinanciamiento / fCredito) * 100
    
End Function

Private Function adeudo() As Double
        
        adeudo = Val(fnstrValor(txttotalpagar.Text)) - Val(fnstrValor(txtpagado.Text)) '- Val(fnstrValor(txtdescuento.Text)) - Val(fnstrValor(txtnotadescuento.Text))

End Function

Private Sub txtcredito_LostFocus()

    If validaForma = True Then
        txtpagodiario.Text = pagoDiario
        txtfinan.Text = financiamiento
        txttotalpagar.Text = totalAPagar
        txtadeudo.Text = adeudo
    End If
    
End Sub

Private Sub txtfinan_GotFocus()
    txtfinan.SelLength = 255
End Sub

Private Sub txtfinan_LostFocus()

    If validaForma = True Then
        txtfinanpor.Text = porcentajeFinanciamiento
        txtpagodiario.Text = pagoDiario
        txttotalpagar.Text = totalAPagar
        txtadeudo.Text = adeudo
    End If
    
End Sub

Private Sub txtfinanpor_GotFocus()
    txtfinanpor.SelLength = 255
End Sub

Private Sub txtfinanpor_LostFocus()
    
    If validaForma = True Then
        txtpagodiario.Text = pagoDiario
        txttotalpagar.Text = totalAPagar
        txtfinan.Text = financiamiento
        txtadeudo.Text = adeudo
    End If
    
End Sub

Private Sub txtfechaini_Change()

    txtfechafin.Caption = Format(Now + CInt(txtpagos.Text), "dd/mm/yyyy")
        
End Sub

Private Sub txtpagos_LostFocus()
    
    If validaForma = True Then
        txtpagodiario.Text = pagoDiario
        txttotalpagar.Text = totalAPagar
        txtfinan.Text = financiamiento
        txtfechafin.Caption = Format(Now + CInt(txtpagos.Text), "dd/mm/yyyy")
        txtadeudo.Text = adeudo
    End If
    
End Sub

Private Sub txtMora_LostFocus()
    
    txttotalpagar.Text = totalAPagar
    txtadeudo.Text = adeudo

End Sub

Private Sub txtnotacred_LostFocus()
    
    txttotalpagar.Text = totalAPagar
    txtadeudo.Text = adeudo
    
End Sub

Private Sub txtdescuento_LostFocus()

    txttotalpagar.Text = totalAPagar
    txtadeudo.Text = adeudo

End Sub

Private Sub txtnotadescuento_LostFocus()

    txttotalpagar.Text = totalAPagar
    txtadeudo.Text = adeudo

End Sub

