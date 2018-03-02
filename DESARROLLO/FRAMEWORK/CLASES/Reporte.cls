VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E73914200F6"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private m_bMuestraBotonCancel As Boolean 'Si es true, muestra el boton de cancelar impresion. Si es False no lo muestra.

Private m_bMuestraBotonesVentana As Boolean 'Si es true muestra los control windos de la venta. Si es False no los muestra.

Private m_bMuestraBotonImpresion As Boolean 'Si es true muestra el boton de impresion. Si es False no lo muestra.

Private m_strNombreReporte As String 'Define el nombre del reporte (formato previamente diseñado)

Private m_bEncontroImpresora As Boolean 'Indica si existe la impresora

Private m_strParametros As Integer

Private m_strFormato As String

Private m_strConsulta As String

Private m_Printer As Printer

Private m_strImpresora As String

Private m_oCrystalReport As CrystalReport

Private m_strParametrosSP As String

Private m_strMensajeFalla As String

Private m_cParametros As Collection

Private m_bVistaPreliminar As Boolean

Private Report As Reporte

Public Property Get cParametros() As Collection
    Set cParametros = m_cParametros
End Property

Public Property Let cParametros(ByVal vData As Collection)
    Set m_cParametros = vData
End Property

Public Property Let strMensajeFalla(ByVal vData As String)
    m_strMensajeFalla = vData
End Property

Public Property Let strParametrosSP(ByVal vData As String)
    m_strParametrosSP = vData
End Property

Public Property Let oCrystalReport(ByVal vData As Object)
    Set m_oCrystalReport = vData
End Property

Public Property Let bMuestraBotonCancel(ByVal vData As Boolean)
    m_oCrystalReport.WindowShowCancelBtn = vData
End Property

Public Property Let bMuestraBotonesVentana(ByVal vData As Boolean)
    m_oCrystalReport.WindowControls = vData
End Property

Public Property Let bVistaPreliminar(ByVal vData As Boolean)
    m_bVistaPreliminar = vData
End Property

Public Property Let bMuestraBotonImpresion(ByVal vData As Boolean)
    m_oCrystalReport.WindowShowPrintBtn = vData
End Property

Public Property Let strNombreReporte(ByVal vData As String)
    m_oCrystalReport.ReportFileName = vData
End Property

Public Property Let strImpresora(ByVal vData As String)
    m_strImpresora = vData
End Property

Public Property Let strConsulta(ByVal vData As String)
    m_strConsulta = vData
End Property

Public Property Get strConsulta() As String
    strConsulta = m_strConsulta
End Property

Public Property Let strFormato(ByVal vData As String)
    m_strFormato = vData
End Property

Public Property Get strFormato() As String
    strFormato = m_strFormato
End Property

Public Function fnImprime() As String
    
    'Verifica que existe la impresora
    m_bEncontroImpresora = False
    If m_bVistaPreliminar = False Then
        Dim tmpPrinter As Printer
        For Each tmpPrinter In Printers
           If tmpPrinter.DeviceName = m_strImpresora Then
              
                Set Printer = tmpPrinter
                m_bEncontroImpresora = True
                Exit For
           End If
        Next
    Else
        m_bEncontroImpresora = True
    End If
    
    If True = m_bEncontroImpresora Then

        'Define si es a pantalla o a impresora
        If m_bVistaPreliminar Then
            m_oCrystalReport.Destination = 0 'crptToWindow (pantalla)
        Else
            m_oCrystalReport.Destination = 1 'crptToPrinter (impresora)
        End If
        
        Dim iParametro As Integer
        Dim oCampo As Campo
        
        'Define los parametros del sp
        iParametro = 0
        For Each oCampo In m_cParametros
            
            m_oCrystalReport.StoredProcParam(iParametro) = oCampo.Valor
            iParametro = iParametro + 1
            
        Next
        
        'Mustra la ventana de vista previa maximizada
        m_oCrystalReport.WindowState = crptMaximized
        
        'Ejecuta la impresión (a pantalla o la impresora)
        m_oCrystalReport.Action = 1
    
    Else
        fnImprime = "La impresora " & m_strImpresora & " no se encuentra activa."
    End If
        
End Function

Private Sub LoadReport()
End Sub

'Dim Report As New CrystalReport2
'Dim m_RS As New ADOR.Recordset        ' Create and ADO record set
'
'Private Sub Form_Load()
'
'    Dim oConeccion As New ADODB.Connection
'    Dim sCnStr As String
'
'    oConeccion.Provider = "sqloledb"
'
'    sCnStr = "Server=ARRIVERO;Database=GOVI;Trusted_Connection=yes"
'
'    oConeccion.Open sCnStr
'
'    ' Open the recordset
'    'm_RS.Open "exec pedidoReporte 42", oConeccion
'
'    Dim ParsFields As CRAXDRT.ParameterFieldDefinitions
'    Dim ParField As CRAXDRT.ParameterFieldDefinition
'
'    Set ParsFields = Report.ParameterFields
'    For Each ParField In ParsFields
'        If ParField.ParameterFieldName = "@iPedidoid" Then
'            ParField.AddCurrentValue 40
'        End If
'    Next
'
'    Report.EnableParameterPrompting = False
'    ' Pass this recordset to the report engine to use as the datasource
'    ' If you comment out the following line the report will show all customer
'    ' in the world because the report was created to show all customers and
'    ' this information is stored in the DSR.
'    'Report.Database.SetDataSource m_RS
'
'    ' The Smart Viewer (CRVIEWER) is the preview window for the report
'    ' If you are don't wish to show the report to the user, ie. only want
'    ' to print the report, then you don't need to use the viewer at all.
'    CRViewer1.ReportSource = Report
'    Report.PrintOut True
'    ' You have full access to all the report objects, so you may want
'    ' to change the text in the title of the report
'    ' Comment out the line below to see the text that is in the report
'    ' definition.
'    'Report.Text6.SetText "You can change text objects through code!"
'
'    '  The Smart viewer has a object model that allows you to modify the
'    '  look and feel of the preview window at runtime.
'    '  The ViewReport method will start the process of the report.
'    CRViewer1.ViewReport
'
'End Sub

