VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form pagosPorCobrador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos por Cobrador"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crpagosPorCobrador 
      Left            =   360
      Top             =   1410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2490
      TabIndex        =   5
      Top             =   1350
      Width           =   1095
   End
   Begin VB.CommandButton cmdreporte 
      Caption         =   "Generar Reporte"
      Height          =   495
      Left            =   1170
      TabIndex        =   4
      Top             =   1350
      Width           =   1095
   End
   Begin VB.ComboBox cbCobrador 
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtfecha 
      Height          =   345
      Left            =   1110
      TabIndex        =   0
      Top             =   300
      Width           =   2385
      _Version        =   65537
      _ExtentX        =   4207
      _ExtentY        =   609
      _StockProps     =   93
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin VB.Label Label8 
      Caption         =   "Cobrador:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   900
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   330
      Width           =   735
   End
End
Attribute VB_Name = "pagosPorCobrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdsalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        fnLlenaComboCollecion cbCobrador, oUsuario.cDatos, 0, ""
    End If
    Set oUsuario = Nothing

End Sub

Private Sub cmdreporte_Click()

    Call imprime("rpPagosPorCobrador", crpagosPorCobrador, txtfecha.Text, cbCobrador.Text)  'cbCobrador.ItemData (cbCobrador.Index)
     
End Sub

Private Sub imprime(strReporte As String, crObjeto As CrystalReport, _
                      strFecha As String, strCobrador As String)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    Dim strNombreReporte As String
    
    If strFecha <> "" Then
    
        cParametros.Add oCampo.CreaCampo(adInteger, , , strFecha)
        cParametros.Add oCampo.CreaCampo(adInteger, , , strCobrador)
        
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

