VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form saldosfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2220
   ClientLeft      =   3285
   ClientTop       =   3735
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crSaldos 
      Left            =   1560
      Top             =   1740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport rep1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\facturas\1.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdir 
      Caption         =   "Ir a reporte"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin EditLib.fpLongInteger txtfolio 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   1185
         _Version        =   196608
         _ExtentX        =   2090
         _ExtentY        =   503
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
         Text            =   "0"
         MaxValue        =   "999999"
         MinValue        =   "1"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
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
      Begin VB.TextBox txtnombre 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox txtapellido 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "No. de Folio:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre(s):"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   750
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido(s):"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1110
         Width           =   1095
      End
   End
End
Attribute VB_Name = "saldosfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iFolio As Long

Private Sub Form_Load()

    txtfolio = iFolio

End Sub


Private Sub txtfolio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
        Dim oCredito As New credito
        Dim cDatos As New Collection
        Dim cRegistro As New Collection
        Dim oCampo As New Campo
        
        Set cDatos = oCredito.obtenGenerales(Val(txtfolio.Text))
        
        If cDatos.Count > 0 Then
        
            Set cRegistro = cDatos(1)
            
            Set oCampo = cRegistro(3)
            txtnombre.Text = oCampo.Valor
            Set oCampo = cRegistro(4)
            txtapellido.Text = oCampo.Valor
            
        Else
            MsgBox "¡El crédito con No. de folio " + txtfolio.Text + "no existe, verifique por favor!", vbInformation + vbOKOnly, "Reportes"
            txtfolio.SetFocus
        End If
        
        Set oCredito = Nothing
            
    End If
    
'    Dim datos As Recordset
'
'    If KeyAscii = 13 Then
'        If IsNumeric(txtfolio.Text) And txtfolio.Text <> "" Then
'            txtfolio.Text = CLng(txtfolio.Text)
'            Set datos = Base.OpenRecordset("select * from ctes_por_credito where factura=" & CStr(txtfolio.Text))
'
'            If datos.RecordCount > 0 Then
'                txtnombre.Text = datos!Nombre
'                txtapellido.Text = datos!apellido
'            Else
'                MsgBox "No existe el folio", vbInformation, "Reportes"
'                txtfolio.Text = ""
'                txtfolio.SetFocus
'            End If
'            datos.Close
'        Else
'            MsgBox "El dato que se introdujo es invalido", vbCritical, "Reportes"
'        End If
'    End If

End Sub

'Private Sub txtnombre_KeyPress(KeyAscii As Integer)
'
'    Dim datos As Recordset
'
'    If KeyAscii = 13 Then
'
'        If txtnombre.Text <> "" Then
'
'            Nombre = Trim(txtnombre.Text)
'
'            If txtnombre.Text = "*" Then
'                Set datos = Base.OpenRecordset("select * from clientes")
'            Else
'                Set datos = Base.OpenRecordset("select * from clientes where Ucase(nombre) like '" & UCase(txtnombre.Text) & "*'")
'            End If
'
'            If datos.RecordCount > 0 Then
'                'nombre = Trim(txtnombre.Text)
'                'band = 2
'                frmlistaclientes1.Show 1
'            Else
'                MsgBox "No existen clientes con ese nombre ", vbInformation, "Reportes"
'                txtnombre.Text = ""
'                txtnombre.SetFocus
'            End If
'
'            datos.Close
'
'        End If
'
'    End If
'
'End Sub

Private Function imprimefn(lFolio As Long)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    cParametros.Add oCampo.CreaCampo(adInteger, , , lFolio)
    
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

End Function

Private Sub cmdir_Click()

    imprimefn Val(txtfolio.Text)
    
'    Unload Me
    
'    Dim datos As Recordset
'    Dim i As Integer
'    i = 0
'    Base.Execute "delete from saldos"
'
'    If txtfolio.Text <> "" Then
'
'        Set datos = Base.OpenRecordset("select * from Saldos_Individuales where factura=" & CStr(txtfolio.Text))
'
'        If datos.RecordCount > 2 Then
'            datos.MoveNext
'            datos.MoveNext
'            While Not datos.EOF
'                i = i + 1
'                Base.Execute "insert into saldos (factura,nombre,Cantidad_Solicitada,Cantidad_Abonada,Dias_de_credito,Saldo_Faltante,Fecha,Cantidad,Status,No_pago) values(" + CStr(datos!factura) + ",'" + datos!Nombre + "'," + CStr(datos!Cantidad_Solicitada) + "," + CStr(datos!Cantidad_Abonada) + "," + CStr(datos!Dias_de_credito) + "," + CStr(datos!Saldo_Faltante) + ",'" + Format(datos!fecha, "dd/mm/yyyy") + "'," + CStr(datos!Cantidad) + ",'" + datos!Status + "'," + CStr(i) + ")"
'                datos.MoveNext
'            Wend
'            datos.MoveFirst
'            datos.Close
'        Else
'            MsgBox "No existe el folio", vbInformation, "Reportes"
'            txtfolio.Text = ""
'            txtfolio.SetFocus
'        End If
'
'        rep1.PrintReport
'        'datos.Close
'    Else
'        txtfolio.SetFocus
'    End If

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

