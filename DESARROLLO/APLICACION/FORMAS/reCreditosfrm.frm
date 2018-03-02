VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form reCreditosfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Créditos"
   ClientHeight    =   1635
   ClientLeft      =   4215
   ClientTop       =   3270
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crCreditosNuevos 
      Left            =   3000
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport renuevo 
      Left            =   120
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Facturas\nuevos.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport retermina 
      Left            =   120
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Facturas\termina.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdreporte 
      Caption         =   "Generar Reporte"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   990
      Width           =   1095
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   990
      Width           =   1095
   End
   Begin SSCalendarWidgets_A.SSDateCombo txtfecha 
      Height          =   345
      Left            =   1380
      TabIndex        =   3
      Top             =   390
      Width           =   1965
      _Version        =   65537
      _ExtentX        =   3466
      _ExtentY        =   609
      _StockProps     =   93
      ScrollBarTracking=   0   'False
      SpinButton      =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   420
      Width           =   735
   End
End
Attribute VB_Name = "reCreditosfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    txtfecha.Text = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub cmdreporte_Click()
    
    'If recredito = 0 Then
    '    reporte_nuevos
    'Else
    '    reporte_termina
    'End If

    imprimefn txtfecha.Text
    
End Sub

Private Function imprimefn(strFecha As String)

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
    
    cParametros.Add oCampo.CreaCampo(adInteger, , , strFecha)
    
    oReporte.oCrystalReport = crCreditosNuevos
    If gstrReporteEnPantalla = "Si" Then
        oReporte.bVistaPreliminar = True
    Else
        oReporte.bVistaPreliminar = False
    End If
    oReporte.strImpresora = gPrintPed
    oReporte.strNombreReporte = DirSys & "creditosNuevos.rpt"
    oReporte.cParametros = cParametros
    oReporte.fnImprime
    Set oReporte = Nothing

End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

'Private Sub reporte_nuevos()
'
'    Dim datos, datos1, datos2 As Recordset
'    Dim pagos1, gastos, folio As Double
'    Dim conta As Integer
'
'    If txtfecha.Text <> "" Then
'
'        Base.Execute "delete from nuevos"
'
'        conta = 0
'        Set datos = Base.OpenRecordset("select * from qrycreditos")
'        While Not datos.EOF
'            If datos!fecha = CDate(txtfecha.Text) Then
'                Base.Execute "insert into nuevos (No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fecha,Status) values(" + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!credito) + ",'" + CStr(datos!Nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "')"
'                conta = 1
'            End If
'            datos.MoveNext
'        Wend
'        datos.Close
'        If conta = 1 Then
'            renuevo.PrintReport
'
'        Else
'            MsgBox "No existen créditos nuevos registrados para el dia seleccionado", vbInformation, "Resumen Diario"
'            txtfecha.SetFocus
'        End If
'    Else
'        MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario"
'        txtfecha.Text = Format(Now, "dd/mm/yyyy")
'        txtfecha.SetFocus
'    End If
'
'End Sub

'Private Sub reporte_termina()
'Dim datos, datos1, datos2 As Recordset
'Dim pagos1, gastos, folio As Double
'Dim conta As Integer
'
'If txtfecha.Text <> "" Then
'
'    Base.Execute "delete from terminados"
'
'    conta = 0
'    Set datos = Base.OpenRecordset("select * from qrycreditosterminados ")
'    While Not datos.EOF
'        If datos!fechatermina = CDate(txtfecha.Text) Then
'            Base.Execute "insert into terminados (No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fecha,Status) values(" + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!credito) + ",'" + CStr(datos!Nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "')"
'            conta = 1
'        End If
'        datos.MoveNext
'    Wend
'    datos.Close
'    If conta = 1 Then
'        retermina.PrintReport
'    Else
'        MsgBox "No existen créditos terminados para el dia seleccionado", vbInformation, "Resumen Diario"
'        txtfecha.SetFocus
'    End If
'
'Else
'    MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario"
'    txtfecha.Text = Format(Now, "dd/mm/yyyy")
'    txtfecha.SetFocus
'End If
'End Sub

