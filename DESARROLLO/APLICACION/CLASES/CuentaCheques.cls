VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentaCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bDatos As Boolean
Private m_cDatos As Collection

Public Property Get cDatos() As Collection
    Set cDatos = m_cDatos
End Property

Public Property Let cDatos(ByRef vData As Collection)
    Set m_cDatos = vData
End Property

Public Function fnCatalogo(iSalon As Integer) As Boolean
    
    Dim m_oBD As New DataBase
    'Obten las entidades
    m_oBD.Procedimiento = "catalogoCuentasCheques"
    m_oBD.Parametros = iSalon
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    fnCatalogo = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function fnCatalogoSaldos(iSalon As Integer) As Boolean
    
    Dim m_oBD As New DataBase
    'Obten las entidades
    m_oBD.Procedimiento = "CUENTA_BANCOSelSaldosCuentasProc"
    m_oBD.Parametros = iSalon
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    fnCatalogoSaldos = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

'Public Function catalogoEsp(iSalon As Integer) As Boolean
Public Function catalogoEsp() As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "catalogoCuentasChequesEsp"
    'm_oBD.Parametros = iSalon
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    catalogoEsp = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

'Public Function nombreCuenta(iSalon As Integer, iCuenta As Integer) As String
Public Function nombreCuenta(iCuenta As Integer) As String
    
    Dim m_oBD As New DataBase
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    'Obten las entidades
    m_oBD.Procedimiento = "CUENTA_BANCOSelNombreCuenta"
    'm_oBD.Parametros = iSalon & ", " & iCuenta
    m_oBD.Parametros = iCuenta
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Set cCheque = m_cDatos(1) 'Numero de cuenta
    Set oCampo = cCheque(1)
    nombreCuenta = oCampo.Valor
    
    Set oCampo = cCheque(2) 'Nombre Cuenta
    nombreCuenta = nombreCuenta & " - " & oCampo.Valor
    
    Set m_oBD = Nothing
    
End Function

'Public Function creaCuenta(iSalon As Integer, cCuenta As Collection) As String
Public Function creaCuenta(cCuenta As Collection) As String

'    Dim oCuenta As New Cuenta
'    Dim strRespuesta As String
    creaCuenta = ""
'    strRespuesta = oAlmacen.contabilidadAbilitada()
'    If "NO" = strRespuesta Then
'        Call oCuenta.obtenConceptos(iSalon, 4)
'        If oCuenta.cDatos.Count <= 0 Then
'            Call registraConceptos
'            Call registraTipoOperacion
'            Call registraOperacionCuenta(iSalon, 0)
'        End If
'    ElseIf "SI_PRIMER_VEZ" = strRespuesta Then
'        'Implica crear catalogo y actualizar relación tipo operació vs cuenta
'        If oCuenta.creaModulo(iSalon) = 0 Then
'            oCuenta.actualizaRelaciones (iSalon)
'        End If
'        creaCuenta = "YA_HAY_MODULO_CONTABLE"
'    End If
'    Set oCuenta = Nothing
           
    Dim m_oBD As New DataBase
    m_oBD.cDetalle = cCuenta
    m_oBD.Procedimiento = "CUENTA_BANCOInsProc"
    m_oBD.Parametros = ""
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsertaDetalle

    Set m_oBD = Nothing
    
End Function

Public Function chequesaActualizaEstatusCobrado(cCheques As Collection)

    Dim m_oBD As New DataBase
    m_oBD.cDetalle = cCheques
    m_oBD.Procedimiento = "CHEQUE_NEGOCIOUpdEstatusProc"
    m_oBD.Parametros = ""
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsertaDetalle

    Set m_oBD = Nothing

End Function

'Public Function chequesInsertaSerie(iSalon As Integer, iCuentaCheques As Integer, iChequeInicial As Integer, iChequeFinal As Integer)
Public Function chequesInsertaSerie(iCuentaCheques As Integer, iChequeInicial As Integer, iChequeFinal As Integer)

    Dim m_oBD As New DataBase
    Dim i As Integer
    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    For i = iChequeInicial To iChequeFinal
    
        Set cCheque = New Collection
        'cCheque.Add oCampo.CreaCampo(adInteger, , , iSalon)
        cCheque.Add oCampo.CreaCampo(adInteger, , , iCuentaCheques)
        cCheque.Add oCampo.CreaCampo(adInteger, , , i)
        cCheques.Add cCheque
            
    Next

    m_oBD.cDetalle = cCheques
    m_oBD.Procedimiento = "CHEQUE_NEGOCIOInsProc"
    m_oBD.Parametros = ""
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsertaDetalle

    Set m_oBD = Nothing

End Function

'Public Function chequesEstatusActual(iSalon As Integer, _
'                                     iCuentaCheques As Integer, _
'                                     iEstadoCheque As Integer) As Boolean
Public Function chequesEstatusActual(iCuentaCheques As Integer, _
                                     iEstadoCheque As Integer) As Boolean

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "CHEQUE_NEGOCIOSelEstatusActual"
    m_oBD.Parametros = iCuentaCheques & ", " & iEstadoCheque
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    chequesEstatusActual = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

'Public Function chequesRegistroActual(iSalon As Integer, _
'                                      iCuentaCheques As Integer, _
'                                      ByRef iChequeInicial As Integer, _
'                                      ByRef iChequeFinal As Integer) As Boolean
Public Function chequesRegistroActual(iCuentaCheques As Integer, _
                                      ByRef iChequeInicial As Integer, _
                                      ByRef iChequeFinal As Integer) As Boolean

    Dim m_oBD As New DataBase
    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    m_oBD.Procedimiento = "CHEQUE_NEGOCIORegActual"
    m_oBD.Parametros = iCuentaCheques
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Set cCheque = m_cDatos.Item(1)
    Set oCampo = cCheque.Item(1)
    iChequeInicial = oCampo.Valor
    
    Set oCampo = cCheque.Item(2)
    iChequeFinal = oCampo.Valor
    
    chequesRegistroActual = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function
'Public Function siguienteDisponible(iSalon As Integer, _
'                                    iCuentaCheques As Integer, _
'                                    ByRef iChequeDisponible As Integer) As Boolean

Public Function siguienteDisponible(iCuentaCheques As Integer, _
                                    ByRef iChequeDisponible As Long) As Boolean

    Dim m_oBD As New DataBase
    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    m_oBD.Procedimiento = "CHEQUE_NEGOCIOSigDisponible"
    m_oBD.Parametros = iCuentaCheques
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    If m_cDatos.Count > 0 Then
        
        Set cCheque = m_cDatos.Item(1)
        Set oCampo = cCheque.Item(1)
        If IsNull(oCampo.Valor) Then
            iChequeDisponible = 0
            siguienteDisponible = False
        Else
            iChequeDisponible = oCampo.Valor
            siguienteDisponible = True
        End If
    Else
        iChequeDisponible = 0
        siguienteDisponible = False
    End If
    
    Set m_oBD = Nothing
    
End Function

'Public Function saldoCuenta(iSalon As Integer, _
'                            iCuentaCheques As Integer, _
'                            ByRef dSaldo As Double) As Boolean
Public Function saldoCuenta(iCuentaCheques As Integer, _
                            ByRef dSaldo As Double) As Boolean

    Dim m_oBD As New DataBase
    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    m_oBD.Procedimiento = "CUENTA_CHEQUESSaldoDisponible"
    m_oBD.Parametros = iCuentaCheques
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Set cCheque = m_cDatos.Item(1)
    Set oCampo = cCheque.Item(1)
    dSaldo = oCampo.Valor
    
    saldoCuenta = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

'Public Function saldoBancos(iSalon As Integer) As Double
Public Function saldoBancos() As Double

    Dim m_oBD As New DataBase
    Dim cCheques As New Collection
    Dim cCheque As Collection
    Dim oCampo As New Campo
    
    m_oBD.Procedimiento = "CUENTA_CHEQUESSaldoBancos"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Set cCheque = m_cDatos.Item(1)
    Set oCampo = cCheque.Item(1)
    If IsNull(oCampo.Valor) Then
        saldoBancos = 0
    Else
        saldoBancos = oCampo.Valor
    End If
    
    Set m_oBD = Nothing
   
End Function

Public Function registraRelacion(cRelaciones As Collection)

    Dim m_oBD As New DataBase
    m_oBD.cDetalle = cRelaciones
    m_oBD.Procedimiento = "CTA_BANCO_CTA_CONTABLEInsProc"
    m_oBD.Parametros = ""
    m_oBD.bDatos = False
    m_oBD.bFolio = False
    m_oBD.fnInsertaDetalle
    Set m_oBD = Nothing
    
End Function

