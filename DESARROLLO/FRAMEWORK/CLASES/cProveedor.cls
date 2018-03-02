VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cDireccion"
Attribute VB_Ext_KEY = "Member1" ,"cFax"
Attribute VB_Ext_KEY = "Member2" ,"cTelefono"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariProveedor As Integer 'local copy
'Public moTelefono As cTelefono
'Public moDireccion As cDireccion
Public mstrGiroComercial As String
Public mstrSituacionActual As String
Public msiEstadoCte As Integer
Public msiPorcentajeDescuento As Integer
Public msiDiasDeCredito As Integer
Public msiEstadoProveedor As Integer
Public msiTipoProveedor As Integer
Public mstrRFC As String
Public mstrMarca As String
Public mstrURL As String
Public mstrRazonSocial As String

Private m_cDirecciones As Collection
Private m_cTelefonos As Collection
Private m_cFaxeS As Collection
Private m_cContactos As Collection

Const DIARIO = 1
Const SEMANAL = 2
Const QUINCENAL = 3
Const MENSUAL = 4
Const BIMESTRAL = 5
Const TRIMESTRAL = 6
Const SEMESTRAL = 7
Const ANUAL = 8

Private m_bDatos As Boolean
Private m_cDatos As Collection
Private m_cDetalle As Collection

Public Property Get cDetalle() As Collection
    Set cDetalle = m_cDetalle
End Property

Public Property Let cDetalle(ByVal vData As Collection)
    Set m_cDetalle = vData
End Property

Public Property Get cDatos() As Collection
    Set cDatos = m_cDatos
End Property

Public Property Let cDatos(ByVal vData As Collection)
    Set m_cDatos = vData
End Property

Public Property Get cDirecciones() As Collection
    Set cDirecciones = m_cDirecciones
End Property

Public Property Let cDirecciones(ByVal vData As Collection)
    Set m_cDirecciones = vData
End Property

Public Property Get cTelefonos() As Collection
    Set cTelefonos = m_cTelefonos
End Property

Public Property Let cTelefonos(ByVal vData As Collection)
    Set m_cTelefonos = vData
End Property

Public Property Get cFaxes() As Collection
    Set cFaxes = m_cFaxeS
End Property

Public Property Let ccFaxes(ByVal vData As Collection)
    Set m_cFaxeS = vData
End Property

Public Property Get cContactos() As Collection
    Set cContactos = m_cContactos
End Property

Public Property Let cContactos(ByVal vData As Collection)
    Set m_cContactos = vData
End Property

Public Property Get bDatos() As Boolean
    bDatos = m_bDatos
End Property

Public Property Let bDatos(ByVal vData As Boolean)
    m_bDatos = vData
End Property

Public Property Let iProveedor(ByVal vData As Integer)
    mvariProveedor = vData
End Property

Public Property Get iProveedor() As Integer
    iProveedor = mvariProveedor
End Property

Public Function Consulta(iSalon As Integer, iProveedor As Integer, iCual As Integer) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PROVEEDORSel"
    m_oBD.Parametros = iSalon & ", " & iProveedor & ", " & iCual
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    If m_cDatos.Count > 0 Then
        
        Dim cRegistro As Collection
        Dim oCampo As Campo
        
        Set cRegistro = m_cDatos(1) 'Obten el registro consultado
        
        Set oCampo = cRegistro(2) 'Id Proveedor
        mvariProveedor = oCampo.Valor
        
        Set oCampo = cRegistro(3) 'Razon Social
        mstrRazonSocial = oCampo.Valor
        
        Set oCampo = cRegistro(4) 'RFC
        mstrRFC = oCampo.Valor
        
        Set oCampo = cRegistro(8) 'Página web
        mstrURL = oCampo.Valor
        
        Set oCampo = cRegistro(6) 'Porcentaje Descuento
        msiPorcentajeDescuento = oCampo.Valor
        
        Set oCampo = cRegistro(7) 'Dias de crédito
        msiDiasDeCredito = oCampo.Valor
        
        Set oCampo = cRegistro(9) 'Giro comercial
        mstrGiroComercial = oCampo.Valor
        
        Set oCampo = cRegistro(11) 'Situación actual
        mstrSituacionActual = oCampo.Valor
        
        Set oCampo = cRegistro(10) 'Estado del proveedor
        msiEstadoProveedor = oCampo.Valor
        
        Set oCampo = cRegistro(5) 'Tipo de proveedor
        msiTipoProveedor = oCampo.Valor
                
        Set m_cDirecciones = obtenDirecciones(iSalon, iProveedor)
        Set m_cTelefonos = obtenTelefonos(iSalon, iProveedor)
        Set m_cFaxeS = obtenFaxes(iSalon, iProveedor)
        Set m_cContactos = obtenContactos(iSalon, iProveedor)
    
    End If
    
    m_bDatos = m_oBD.bDatos
    Consulta = m_bDatos
    
    Set m_oBD = Nothing
    
End Function

Private Function obtenDirecciones(iSalon As Integer, iProveedor As Integer) As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "DIRECCION_PROVEEDORSel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    
    Set obtenDirecciones = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

Private Function obtenTelefonos(iSalon As Integer, iProveedor As Integer) As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "TELEFONO_PROVEEDORSel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    
    Set obtenTelefonos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

Private Function obtenFaxes(iSalon As Integer, iProveedor As Integer) As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "FAX_PROVEEDORSel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    
    Set obtenFaxes = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

Private Function obtenContactos(iSalon As Integer, iProveedor As Integer) As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CONTACTO_PROVEEDORSel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    
    Set obtenContactos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function fnConsultaCatalogo(iSalon As Integer) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PROVEEDORCatSel"
    m_oBD.Parametros = iSalon
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    fnConsultaCatalogo = m_bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function periodicidad() As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PERIODOSel"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    periodicidad = m_bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function formas() As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "FORMA_PAGOSelCatalogo"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    formas = m_bDatos
    Set m_oBD = Nothing

End Function

Public Function buscaConNombre(iSalon As Integer, strConcepto As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "buscaNombreConcepto"
    m_oBD.Parametros = iSalon & ", '" & strConcepto & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    buscaConNombre = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

'Public Function creaNuevo(iSalon As Integer, _
'                          cPago As Collection, _
'                          dMonto As Double, _
'                          dtProximoPago As String, _
'                          dtUltimoPago As String, _
'                          iPeriodicidad As Integer, _
'                          iNumeroPagosRestantes As Integer, _
'                          strFolioDocumento As String) As Integer
Public Function creaNuevo(cPago As Collection, _
                          dMonto As Double, _
                          dtProximoPago As String, _
                          dtUltimoPago As String, _
                          iPeriodicidad As Integer, _
                          iNumeroPagosRestantes As Integer, _
                          strFolioDocumento As String) As Integer
    
    Dim iPago As Integer
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PROYECCION_EGRESOSInsProc"
    m_oBD.Parametros = ""
    m_oBD.bFolio = True
    m_oBD.cDetalle = cPago
    iPago = m_oBD.fnInsert
    
    If iPago > 0 Then
        
        Dim cPagos As New Collection
        Dim cRegistro As Collection
        Dim oCampo As New Campo
        Dim iIntervalo As Integer
        Dim strPeriodo As String
        Dim iPeriodo As Integer
        
        Select Case iPeriodicidad
            Case Is = DIARIO
                strPeriodo = "d"
                iPeriodo = 1
            Case Is = SEMANAL
                strPeriodo = "ww"
                iPeriodo = 1
            Case Is = QUINCENAL
                strPeriodo = "ww"
                iPeriodo = 2
            Case Is = MENSUAL
                strPeriodo = "m"
                iPeriodo = 1
            Case Is = BIMESTRAL
                strPeriodo = "m"
                iPeriodo = 2
            Case Is = TRIMESTRAL
                strPeriodo = "m"
                iPeriodo = 3
            Case Is = SEMESTRAL
                strPeriodo = "m"
                iPeriodo = 6
            Case Is = ANUAL
                strPeriodo = "yyyy"
                iPeriodo = 1
        End Select
        
        If iNumeroPagosRestantes = 0 Then
            If DateDiff(strPeriodo, dtProximoPago, dtUltimoPago) > 0 Then
                iNumeroPagosRestantes = DateDiff(strPeriodo, dtProximoPago, dtUltimoPago)
            Else
                iNumeroPagosRestantes = 1
            End If
        End If
        
        For iIntervalo = 1 To iNumeroPagosRestantes Step iPeriodo
            
            Set cRegistro = New Collection
            'cRegistro.Add oCampo.CreaCampo(adInteger, , , iSalon)
            cRegistro.Add oCampo.CreaCampo(adInteger, , , dtProximoPago)
            cRegistro.Add oCampo.CreaCampo(adInteger, , , dMonto)
            cRegistro.Add oCampo.CreaCampo(adInteger, , , iPago)
            cRegistro.Add oCampo.CreaCampo(adInteger, , , strFolioDocumento)
            
            dtProximoPago = DateAdd(strPeriodo, iPeriodo, dtProximoPago)
            
            cPagos.Add cRegistro
        Next iIntervalo
        
        m_oBD.cDetalle = cPagos
        m_oBD.Procedimiento = "FRECUENCIAInsProc"
        m_oBD.Parametros = ""
        m_oBD.bDatos = False
        m_oBD.bFolio = False
        m_oBD.fnInsertaDetalle
        
    End If
    
    creaNuevo = iPago
    Set m_oBD = Nothing
    
End Function

Public Function periodicidadDias() As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PERIODODiasSel"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    periodicidadDias = m_bDatos
    Set m_oBD = Nothing

End Function

'Public Function obtenPagos(iSalon As Integer, dtFechaInicial As String, dtFechaFinal As String) As Boolean
Public Function obtenPagos(dtFechaInicial As String, dtFechaFinal As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "pagosProc"
    'm_oBD.Parametros = iSalon & ", '" & dtFechaInicial & "', '" & dtFechaFinal & "'"
    m_oBD.Parametros = "'" & dtFechaInicial & "', '" & dtFechaFinal & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    obtenPagos = m_bDatos
    Set m_oBD = Nothing

End Function

'Public Function obtenPagosVencidos(iSalon As Integer, dtFechaInicial As String) As Boolean
Public Function obtenPagosVencidos(dtFechaInicial As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "pagosVencidosProc"
    'm_oBD.Parametros = iSalon & ", '" & dtFechaInicial & "'"
    m_oBD.Parametros = "'" & dtFechaInicial & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    obtenPagosVencidos = m_bDatos
    Set m_oBD = Nothing

End Function

Public Function pagosRealizados(iSalon As Integer, dtFechaInicial As String, dtFechaFinal As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "pagosRealizados"
    m_oBD.Parametros = iSalon & ", '" & dtFechaInicial & "', '" & dtFechaFinal & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    pagosRealizados = m_bDatos
    Set m_oBD = Nothing

End Function

'Public Function pagoRegistra(iSalon As Integer, cPagos As Collection, cCheques As Collection, iConcepto As Integer, strConcepto As String) As String
Public Function pagoRegistra(cPagos As Collection, cCheques As Collection, iConcepto As Integer, strConcepto As String) As String

'    If oAlmacen.contabilidadAbilitada = "SI" Then
'
'        Dim cRegistro As Collection
'        Dim oCuenta As New Cuenta
'        Dim iCheque As Integer
'        Dim iCuentaBanco As Integer
'        Dim iCuentaContable As Integer
'        Dim iCuentaContableCtaCheques As Integer
'        Dim iCuentaContableIvaAcreditable As Integer
'        Dim iMovimientoContable As Integer
'
        Dim m_oBD As New DataBase
'        Dim oCampo As New Campo
'
'        pagoRegistra = ""
'        iCuentaContable = oCuenta.obtenCuenta(gAlmacen, iConcepto)
'        If iCuentaContable > 0 Then
'
'            iCuentaContableIvaAcreditable = oCuenta.obtenCuentaIvaAcreditable(iSalon)
'
'            If iCuentaContableIvaAcreditable > 0 Then
'
'                'Asiento(s) contable(s) de la(s) cuenta(s) de cheques
'                For iCheque = 1 To cCheques.Count
'
'                    Set cRegistro = cCheques(iCheque)
'
'                    Set oCampo = cRegistro(2)
'                    iCuentaBanco = oCampo.Valor
'                    iCuentaContableCtaCheques = oCuenta.cuentaContableBanco(iSalon, iCuentaBanco)
'
'                    If iCuentaContableCtaCheques <= 0 Then
'                        pagoRegistra = "NO_EXISTE_CUENTA_CONTABLE_CHEQUES"
'                        Exit For
'                    End If
'                Next iCheque
'
'                If pagoRegistra <> "" Then
'                    GoTo fin
'                End If
'            Else
'                pagoRegistra = "NO_EXISTE_CUENTA_IVA_ACREDITABLE"
'                GoTo fin
'            End If
'        Else
'            pagoRegistra = "NO_EXISTE_CUENTA_CONCEPTO"
'            GoTo fin
'        End If
'
'        iMovimientoContable = oCuenta.obtenSiguienteMovimiento(iSalon)
'
'        'REGISTRA EL PAGO
'        m_oBD.Procedimiento = "pagoRegistra"
'        m_oBD.Parametros = ""
'        m_oBD.cDetalle = cPagos
'        Call m_oBD.fnInsertaDetalle
'
'        'REGISTRA EL O MOVIMIENTOS EN CUENTA(S) DE CHEQUES CON QUE SE CUBRE EL PAGO
'        m_oBD.Procedimiento = "MOVIMIENTO_CUENTA_BANCORegistra"
'        m_oBD.cDetalle = cCheques
'        Call m_oBD.fnInsertaDetalle
'
'        Set m_oBD = Nothing
'
'        'REALIZA LOS ASIENTOS CONTABLES.
'        Dim cAsientos As New Collection
'        Dim strFechaPago As String
'        Dim dMontoPago As Double
'        Dim dIVA As Double
'        Dim dMontoIVA As Double
'        Dim dMontoCheque As Double
'
'        Dim cAsiento As New Collection
'
'        'Asiento contable de la cuenta de gastos
'        Set cRegistro = cPagos(1)
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , iSalon)
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , iCuentaContable)
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , iMovimientoContable)
'        Set oCampo = cRegistro(6)
'        strFechaPago = oCampo.Valor
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , strFechaPago) 'fecha pago (movimiento)
'        Set oCampo = cRegistro(5)
'        dMontoPago = oCampo.Valor
'
'        dIVA = oAlmacen.obtenIVA
'        dMontoIVA = dMontoPago * (dIVA / 100#)
'        dMontoPago = dMontoPago - dMontoIVA
'
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , dMontoPago) 'Abono
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , 0) 'Cargo
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , strConcepto) 'Concepto
'        cAsientos.Add cAsiento
'
'        'Asiento contable del IVA acreditable
'        Set cAsiento = New Collection
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , iSalon)
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , iCuentaContableIvaAcreditable) 'id de la cuenta contable
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , iMovimientoContable)
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , strFechaPago)
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , dMontoIVA) 'Abono
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , 0) 'Cargo
'        cAsiento.Add oCampo.CreaCampo(adInteger, , , strConcepto) 'Concepto
'        cAsientos.Add cAsiento
'
'        'Asiento(s) contable(s) de la(s) cuenta(s) de cheques
'        For iCheque = 1 To cCheques.Count
'
'            Set cAsiento = New Collection
'            Set cRegistro = cCheques(iCheque)
'
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , iSalon)
'            Set oCampo = cRegistro(2)
'            iCuentaBanco = oCampo.Valor
'            iCuentaContableCtaCheques = oCuenta.cuentaContableBanco(iSalon, iCuentaBanco)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , iCuentaContableCtaCheques)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , iMovimientoContable)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , strFechaPago)
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , 0) 'Abono
'            Set oCampo = cRegistro(4)
'            dMontoCheque = oCampo.Valor
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , dMontoCheque) 'Cargo
'            cAsiento.Add oCampo.CreaCampo(adInteger, , , strConcepto) 'Concepto
'
'            cAsientos.Add cAsiento
'
'        Next iCheque
'
'        Call oCuenta.registraAsiento(cAsientos)
'
'        Set oCuenta = Nothing
'
'    Else
    
        'REGISTRA EL PAGO
        m_oBD.Procedimiento = "pagoRegistra"
        m_oBD.Parametros = ""
        m_oBD.cDetalle = cPagos
        Call m_oBD.fnInsertaDetalle
        
        'REGISTRA EL O MOVIMIENTOS EN CUENTA(S) DE CHEQUES CON QUE SE CUBRE EL PAGO
        m_oBD.Procedimiento = "MOVIMIENTO_CUENTA_BANCORegistra"
        m_oBD.cDetalle = cCheques
        Call m_oBD.fnInsertaDetalle
        
        Set m_oBD = Nothing
        
'    End If
    
'fin:

End Function

Public Function pagoRegistraProrroga(cParciales As Collection)

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "FRECUENCIAInsProc"
    m_oBD.Parametros = ""
    m_oBD.cDetalle = cParciales
    
    Call m_oBD.fnInsertaDetalle
    Set m_oBD = Nothing
    
End Function

Public Function obtenCuentaContable(iSalon As Integer, iConcepto As Integer) As Integer

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "OPERACION_CUENTAObtenCuenta"
    m_oBD.Parametros = iSalon & ", " & iConcepto
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    Set cRegistro = m_cDatos(1)
    Set oCampo = cRegistro(1)
    If IsNull(oCampo.Valor) Then
        obtenCuentaContable = 0
    Else
        obtenCuentaContable = oCampo.Valor
    End If
    Set m_oBD = Nothing
End Function

Public Function catalogoGastos(iTipoOperacion As Integer) As Boolean

    Dim oBD As New DataBase
    oBD.Procedimiento = "OPERACION_CUENTASelGastos"
    'oBD.Parametros = gAlmacen & ", " & iTipoOperacion
    oBD.Parametros = iTipoOperacion
    
    Set m_cDatos = oBD.fnObtenInformacion()
    m_bDatos = oBD.bDatos
    catalogoGastos = m_bDatos
    
    Set oBD = Nothing
    
End Function

'Se ejecuta, por que no hay un catálgo de gastos, es probable que esté habilitado el módulo contable y se reguiera creare el modelo contable.
Public Function creaModulo(iSalon As Integer)

    'If oAlmacen.contabilidadAbilitada() = "NO" Then
        Call registraConceptos
        Call registraTipoOperacion
        Call SICRED.registraOperacionCuenta(iSalon)
    'Else
    '    Dim oCuenta As New Cuenta
    '
    '    oCuenta.creaModulo (iSalon)
    '
    '    Set oCuenta = Nothing
    'End If
    
End Function

Public Function obtenEstatus() As Boolean

    Dim oBD As New DataBase
    oBD.Procedimiento = "ESTATUS_PROVEEDORSel"
    oBD.Parametros = ""
    
    Set m_cDatos = oBD.fnObtenInformacion()
    m_bDatos = oBD.bDatos
    obtenEstatus = m_bDatos
    
    Set oBD = Nothing
    
End Function

Public Function obtenCatTipo() As Boolean

    Dim oBD As New DataBase
    oBD.Procedimiento = "TIPO_PROVEEDORSel"
    oBD.Parametros = ""
    
    Set m_cDatos = oBD.fnObtenInformacion()
    m_bDatos = oBD.bDatos
    obtenCatTipo = m_bDatos
    
    Set oBD = Nothing
    
End Function

Public Function registra(iSalon As Integer, ByRef iContacto As Integer, _
                          cContacto As Collection, _
                          cDirecciones As Collection, _
                          cTelefonos As Collection, _
                          cFaxes As Collection, _
                          cContactos As Collection) As Boolean
    
    Dim m_oBD As New DataBase
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
                      
    If iContacto <= 0 Then 'Crear nuevo contacto
        
        m_oBD.Procedimiento = "PROVEEDORIns"
        m_oBD.Parametros = ""
        m_oBD.bFolio = True
        m_oBD.cDetalle = cContacto
        iContacto = m_oBD.fnInsert
        
        'actualiza Direcciones
        If cDirecciones.Count > 0 Then
            For Each cRegistro In cDirecciones
                Set oCampo = cRegistro(2)
                oCampo.Valor = iProveedor
            Next cRegistro
            Call registraDirecciones(iSalon, iProveedor, cDirecciones)
        End If
        
        'actualiza telefonos
        If cTelefonos.Count > 0 Then
            For Each cRegistro In cTelefonos
                Set oCampo = cRegistro(2)
                oCampo.Valor = iContacto
            Next cRegistro
            Call registraTelefonos(iSalon, iProveedor, cTelefonos)
        End If
    
        'actualiza FAXES
        If cFaxes.Count > 0 Then
            For Each cRegistro In cFaxes
                Set oCampo = cRegistro(2)
                oCampo.Valor = iProveedor
            Next cRegistro
            Call registraFaxes(iSalon, iProveedor, cFaxes)
        End If
    
        'actualiza CONTCATOS
        If cContactos.Count > 0 Then
            For Each cRegistro In cContactos
                Set oCampo = cRegistro(2)
                oCampo.Valor = iProveedor
            Next cRegistro
            Call registraContactos(iSalon, iProveedor, cContactos)
        End If
    
    Else    'Actualiza el proveedor
    
        m_oBD.Procedimiento = "PROVEEDORUpd"
        m_oBD.Parametros = iSalon & ", " & iProveedor
        m_oBD.bFolio = False
        m_oBD.cDetalle = cContacto
        m_oBD.fnInsertaDetalle
        
        If cDirecciones.Count > 0 Then
            Call registraDirecciones(iSalon, iProveedor, cDirecciones)
        End If
        
        If cTelefonos.Count > 0 Then
            Call registraTelefonos(iSalon, iContacto, cTelefonos)
        End If
            
        If cFaxes.Count > 0 Then
            Call registraFaxes(iSalon, iProveedor, cFaxes)
        End If
            
        If cContactos.Count > 0 Then
            Call registraContactos(iSalon, iProveedor, cContactos)
        End If
        
    End If
    
    Set m_oBD = Nothing
    
End Function

Private Function registraDirecciones(iSalon As Integer, _
                                    iProveedor As Integer, _
                                    cDirecciones As Collection)
    
    Dim m_oBD As New DataBase
    'Elimina los que hay
    m_oBD.Procedimiento = "DIRECCION_PROVEEDORDel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsert
    
    'registra los nuevos
    If cDirecciones.Count > 0 Then
        m_oBD.cDetalle = cDirecciones
        m_oBD.Procedimiento = "DIRECCION_PROVEEDORIns"
        m_oBD.Parametros = ""
        m_oBD.fnInsertaDetalle
    End If
    Set m_oBD = Nothing
    
End Function

Private Function registraTelefonos(iSalon As Integer, _
                                  iProveedor As Integer, _
                                  cTelefomos As Collection)
    Dim m_oBD As New DataBase
    
    'Elimina los que hay
    m_oBD.Procedimiento = "TELEFONO_PROVEEDORDel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsert
    
    'registra los nuevos
    If cTelefomos.Count > 0 Then
        m_oBD.cDetalle = cTelefomos
        m_oBD.Procedimiento = "TELEFONO_PROVEEDORIns"
        m_oBD.Parametros = ""
        m_oBD.fnInsertaDetalle
    End If
    
    Set m_oBD = Nothing
End Function

Private Function registraFaxes(iSalon As Integer, _
                                         iProveedor As Integer, _
                                         cIdentificaciones As Collection)
    
    Dim m_oBD As New DataBase
    'Elimina los que hay
    m_oBD.Procedimiento = "FAX_PROVEEDORDel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsert
    
    'registra los nuevos
    If cIdentificaciones.Count > 0 Then
        m_oBD.cDetalle = cIdentificaciones
        m_oBD.Procedimiento = "FAX_PROVEEDORIns"
        m_oBD.Parametros = ""
        m_oBD.fnInsertaDetalle
    End If
    Set m_oBD = Nothing
    
End Function

Private Function registraContactos(iSalon As Integer, _
                                iProveedor As Integer, _
                                cCorreos As Collection)
    Dim m_oBD As New DataBase
    
    'Elimina los que hay
    m_oBD.Procedimiento = "CONTACTO_PROVEEDORDel"
    m_oBD.Parametros = iSalon & ", " & iProveedor
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsert
    
    'registra los nuevos
    If cCorreos.Count > 0 Then
        m_oBD.cDetalle = cCorreos
        m_oBD.Procedimiento = "CONTACTO_PROVEEDORInsProc"
        m_oBD.Parametros = ""
        m_oBD.fnInsertaDetalle
    End If
    Set m_oBD = Nothing
End Function

Public Function registraConcepto(cConceptos As Collection)

    Dim oBD As New DataBase
    oBD.Procedimiento = "CONCEPTOInsProc"
    oBD.cDetalle = cConceptos
    
    oBD.fnInsertaDetalle
    
    Set oBD = Nothing
    
End Function

Public Function registraOperacion(cOperacion As Collection)

    Dim oBD As New DataBase
    oBD.Procedimiento = "OPERACIONInsProc"
    oBD.cDetalle = cOperacion
    
    oBD.fnInsertaDetalle
    
    Set oBD = Nothing
    
End Function

Public Function registraOperacionCuenta(cOperaciones As Collection)

    Dim oBD As New DataBase
    oBD.Procedimiento = "OPERACION_CUENTAInsProc"
    oBD.cDetalle = cOperaciones
    
    oBD.fnInsertaDetalle
    
    Set oBD = Nothing
    
End Function

