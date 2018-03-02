VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form repPagosUsuariofrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Pagos por Usuario"
   ClientHeight    =   2910
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crPagosUsuario 
      Left            =   90
      Top             =   2190
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport repaux 
      Left            =   3120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Facturas\reppagos3.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdreporte 
      Caption         =   "Generar Reporte"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmbUsuario 
      Height          =   315
      ItemData        =   "repPagosUsuariofrm.frx":0000
      Left            =   1440
      List            =   "repPagosUsuariofrm.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtFechaIni 
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   1020
      Width           =   1965
      _Version        =   65537
      _ExtentX        =   3466
      _ExtentY        =   609
      _StockProps     =   93
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtFechaFin 
      Height          =   345
      Left            =   1440
      TabIndex        =   7
      Top             =   1500
      Width           =   1965
      _Version        =   65537
      _ExtentX        =   3466
      _ExtentY        =   609
      _StockProps     =   93
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Final:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicial:"
      Height          =   255
      Left            =   420
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbUsuario 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   780
      TabIndex        =   1
      Top             =   510
      Width           =   615
   End
End
Attribute VB_Name = "repPagosUsuariofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        fnLlenaComboCollecion cmbUsuario, oUsuario.cDatos, 1, "Todos"
        cmbUsuario.ListIndex = 0
    End If
    Set oUsuario = Nothing
    
'    Dim datos As Recordset
'    Dim fecha As Date
'    Dim conta As Integer
'
'    cmbUsuario.AddItem ("Todos")
'    Set datos = Base.OpenRecordset("select distinct usuario from pagos")
'    While Not datos.EOF
'        cmbUsuario.AddItem (datos!usuario)
'        datos.MoveNext
'    Wend
'
'    cmbUsuario.RemoveItem (1)
'    txtFechaIni.Text = Format(Now, "dd/mm/yyyy")
'    txtFechaFin.Text = Format(Now, "dd/mm/yyyy")
'
'    datos.Close

End Sub

Private Function validaForma() As Boolean
    
    Dim strMensaje As String
    Dim iFechaErronea As Integer
    Dim bCorrecta As Boolean
    
    bCorrecta = True
    
    If DateDiff("d", txtFechaFin.Text, txtFechaIni.Text) > 0 Then
        strMensaje = "La fecha final no puede ser menor o igual a la fecha inicial."
        MsgBox strMensaje, vbInformation + vbOKOnly
        txtFechaFin.SetFocus
        bCorrecta = False
    End If
    
    validaForma = bCorrecta
    
End Function

Private Function imprimefn(strFechaIni As String, strFechaFin As String, strCobrador As String)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    cParametros.Add oCampo.CreaCampo(adInteger, , , strFechaIni)
    cParametros.Add oCampo.CreaCampo(adInteger, , , strFechaFin)
    cParametros.Add oCampo.CreaCampo(adInteger, , , strCobrador)
    
    oReporte.oCrystalReport = crPagosUsuario
    If gstrReporteEnPantalla = "Si" Then
        oReporte.bVistaPreliminar = True
    Else
        oReporte.bVistaPreliminar = False
    End If
    oReporte.strImpresora = gPrintPed
    oReporte.strNombreReporte = DirSys & "pagosUsuario.rpt"
    oReporte.cParametros = cParametros
    oReporte.fnImprime
    Set oReporte = Nothing

End Function

Private Sub cmdreporte_Click()

    If validaForma() = False Then
        Exit Sub
    End If
    
    Call imprimefn(txtFechaIni.Text, txtFechaFin.Text, cmbUsuario.Text)



'    If cmbUsuario.Text = "" Then
'    MsgBox ("El campo de usuario no puede estar vacio")
'    Exit Sub
'    End If
'
'    If txtFechaIni.Text = "" Then
'    MsgBox ("El campo de fecha inicial no puede estar vacio")
'    Exit Sub
'    End If
'
'    If txtFechaFin.Text = "" Then
'    MsgBox ("El campo de fecha final no puede estar vacio")
'    Exit Sub
'    End If
    
'    Dim fechaini As Date
'    Dim fechafin As Date
'    Dim datos As Recordset
'    Dim nombre As String
'
'    fechaini = CDate(txtFechaIni.Text)
'    fechafin = CDate(txtFechaFin.Text)
'
'    If cmbUsuario.Text = "Todos" Then
'        Set datos = Base.OpenRecordset("Select *, pagos.no_cliente as no_cliente from pagos,clientes where pagos.no_cliente = clientes.no_cliente and cdbl(fecha) between " + CStr(CDbl(fechaini)) + " and " + CStr(CDbl(fechafin)) + " order by fecha,usuario,hora")
'    Else
'        Set datos = Base.OpenRecordset("Select clientes.nombre, clientes.apellido, pagos.factura,pagos.no_cliente,pagos.cons_pago,pagos.cantpagada,pagos.usuario,pagos.fecha,pagos.hora,pagos.lugar,pagos.cantadeudada,pagos.orden from pagos,clientes where pagos.no_cliente = clientes.no_cliente and usuario = '" + cmbUsuario.Text + "' and fecha between " + CStr(CDbl(fechaini)) + " and " + CStr(CDbl(fechafin)) + " order by fecha,hora")
'    End If
'
'    Base.Execute "delete from pagos2_temp"
'
'    While Not datos.EOF
'
'        Dim nocliente, folio, no_pago, consecutivo, orden As Long
'        Dim fecha As Date
'        Dim pago, adeudo As Double
'        Dim usuario, hora, lugar As String
'
'        nombre = CStr(datos!nombre) + " " + CStr(datos!apellido)
'        folio = CLng(datos!factura)
'        nocliente = CLng(datos!no_cliente)
'        consecutivo = CLng(datos!cons_pago)
'        pago = CDbl(datos!Cantpagada)
'        usuario = CStr(datos!usuario)
'        fecha = CDate(datos!fecha)
'        hora = CStr(datos!hora)
'        lugar = CStr(datos!lugar)
'        adeudo = CDbl(datos!Cantadeudada)
'        orden = CLng(datos!orden)
'
'        Base.Execute "insert into pagos2_temp (nombre,no_cliente,factura,fecha,Cantpagada,Cantadeudada,orden,cons_pago,usuario,hora, lugar ) values('" + CStr(nombre) + "'," + CStr(nocliente) + "," + CStr(folio) + ",'" + CStr(fecha) + "'," + CStr(pago) + "," + CStr(adeudo) + "," + CStr(orden) + "," + CStr(consecutivo) + ",'" + CStr(usuario) + "','" + CStr(hora) + " ','" + CStr(lugar) + "')"
'
'        datos.MoveNext
'
'    Wend
'
'    datos.Close
'
'    repaux.PrintReport
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

