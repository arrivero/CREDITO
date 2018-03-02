VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Credito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bDatos As Boolean
Private m_cDatos As Collection
Private m_cDetalle As Collection

Private m_iFactura As Long

Public Property Get cDatos() As Collection
    Set cDatos = m_cDatos
End Property

Public Property Let cDatos(ByVal vData As Collection)
    Set m_cDatos = vData
End Property

Public Property Get bDatos() As Boolean
    bDatos = m_bDatos
End Property

Public Property Let bDatos(ByVal vData As Boolean)
    m_bDatos = vData
End Property

Public Property Get cDetalle() As Collection
    Set cDetalle = m_cDetalle
End Property

Public Property Let cDetalle(ByVal vData As Collection)
    Set m_cDetalle = vData
End Property

Public Property Get iFactura() As Boolean
    iFactura = m_iFactura
End Property

Public Property Let iFactura(ByVal vData As Boolean)
    m_iFactura = vData
End Property

Public Function registraCredito(iNoCliente As Integer, fCredito As Double, fPagoDiario As Double, fFinanciamiento As Double, _
                                fTotalPagar As Double, iNoPagos As Integer, strFechaInicial As String, strFechaFinal As String, _
                                strFechaContrato As String, strStatus As String, strNoCheque As String, strCobrador As String, _
                                iElectrico As Integer, strDescProducto As String) As Long

    Dim m_oBD As New DataBase

    Dim cCredito As New Collection
    Dim cPagos As New Collection
    Dim cRegistro As New Collection
    Dim cRegistroPago As Collection
    Dim oCampo As New Campo
    Dim iNoPago As Integer
    Dim dFechaPago As Date
    
    cRegistro.Add oCampo.CreaCampo(adInteger, , , iNoCliente)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , fCredito)   'Monto otorgado en el crèdito
    cRegistro.Add oCampo.CreaCampo(adInteger, , , fPagoDiario) 'Cantidad a pagar diario
    cRegistro.Add oCampo.CreaCampo(adInteger, , , fFinanciamiento) 'Total a pagar de financiamiento
    cRegistro.Add oCampo.CreaCampo(adInteger, , , fTotalPagar) 'Cantidad total a pagar
    cRegistro.Add oCampo.CreaCampo(adInteger, , , iNoPagos) 'Numero de pagos
    cRegistro.Add oCampo.CreaCampo(adInteger, , , strFechaInicial)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , strFechaFinal)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , strFechaContrato)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , strStatus)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , iElectrico)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , strDescProducto)
    cRegistro.Add oCampo.CreaCampo(adInteger, , , convierteMontoConLentra(CStr(fCredito))) 'Monto de credito otorgado descrito en letra
    cRegistro.Add oCampo.CreaCampo(adInteger, , , convierteMontoConLentra(CStr(fTotalPagar))) 'Monto total a pagar del credito otorgado, descrito en letra
    'cRegistro.Add oCampo.CreaCampo(adInteger, , , convierteCantidadConLetra(CStr(fCredito))) 'Monto de credito otorgado descrito en letra
    'cRegistro.Add oCampo.CreaCampo(adInteger, , , convierteCantidadConLetra(CStr(fTotalPagar))) 'Monto total a pagar del credito otorgado, descrito en letra
    cRegistro.Add oCampo.CreaCampo(adInteger, , , strNoCheque)
    
    cCredito.Add cRegistro
    
    m_oBD.Procedimiento = "CREDITOAlta"
    m_oBD.Parametros = ""
    m_oBD.bFolio = True
    m_oBD.cDetalle = cCredito
    m_iFactura = m_oBD.fnInsert
    
    If m_iFactura > 0 Then
    
        For iNoPago = 1 To iNoPagos
        
            Set cRegistroPago = New Collection
            'dFechaPago = DateAdd("d", iNoPago, strFechaContrato)
            dFechaPago = DateAdd("d", iNoPago - 1, strFechaInicial)
            
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , m_iFactura)   'Factura
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , 1)   'No de Pago para HH
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , 0) 'Consecutivo de pago
            'cRegistroPago.Add oCampo.CreaCampo(adInteger, , , fPagoDiario) 'Pago (abono)
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , 0#)  'Pago (abono)
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , strCobrador) 'Usuario
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , dFechaPago) 'Fecha de pago
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , "") 'Hora
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , "") 'Lugar
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , iNoPago) 'orden
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , 1) 'indica que debe grabar
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , iNoCliente) 'Cliente
            cRegistroPago.Add oCampo.CreaCampo(adInteger, , , 0#)  'Adeudo
            
            cPagos.Add cRegistroPago
            
        Next iNoPago
        
        m_oBD.Procedimiento = "PAGOGraba"
        m_oBD.Parametros = ""
        m_oBD.bFolio = False
        m_oBD.cDetalle = cPagos
        m_oBD.fnInsertaDetalle
        
    End If
    
    
    registraCredito = m_iFactura
    Set m_oBD = Nothing

    'Los datos del credito son los siguientes (respetar el orden):
    
    'no_cliente int
    'factura int
    'credito float
    'cantpagar float
    'financiamiento float
    'canttotal float
    'no_pagos int
    'fechaini samalldatetime
    'fechatermina samalldatetime
    'fecha samalldatetime
    'Status varchar(5)
    'electrico int
    'descripcion varchar(50)
    'iTipoCredito => segun el valor (1 ó 0) se registra en base de datos; si es uno se ingresa el contenido del dato descripcion
                                                                          'si es cero no se ingresa valor alguno al dato descripcion.
    
'COnsideraciones:
   'iTipoCredito puede ser 1 ó 0 y se registra en base de datos
        'Si es uno se ingresa el valor del campo 'descripcion' en el campo descripcion
        'Si es cero no se ingresa valor alguno al campo descripcion.
     
    'factura, este es un consecutivo, se tiene que buscar la ultima factura y agregar 1, para tener la nueva factura.
        'Ejemplo, en base de datos la última factura es la 234, por tanto la nueva debe ser 235.
        'En la clase DataBase esta la función fnInser, si se usa esta con la propiedad m_bFolio a true te crea el nuevo registro 'padre'
        'El valor de regreso debe ser el número de la nueva factura, pero si falló la inserción, debe regresar un 0 (cero)
        
    'el nuevo credito se ingresa a la entidad 'creditos'
    
    'Ver definiciòn de tabla creditos incluida en el correo
    
End Function

Public Function siguiente() As Long
    
    Dim cRegistros As New Collection
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOSiguiente"
    m_oBD.bFolio = False
    Set cRegistros = m_oBD.fnObtenInformacion
    Set m_oBD = Nothing
    
    If cRegistros.Count > 0 Then
        Set cRegistro = cRegistros(1)
        Set oCampo = cRegistro(1)
        
        siguiente = oCampo.Valor
    Else
        siguiente = 0
    End If
    
End Function

Public Function actualizaCredito(cCreditoDatos As Collection) As Integer

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOActualiza"
    m_oBD.bFolio = False
    m_oBD.cDetalle = cCreditoDatos
    m_oBD.fnInsertaDetalle
    Set m_oBD = Nothing

    'Los datos del credito son los siguientes (respetar el orden):
    
    'no_cliente
    'factura
    'cantpagar
    'financiamiento
    'no_pagos
    'Status
    'iTipoCredito

'COnsideraciones:
   'iTipoCredito puede ser 1 ó 0 y se registra en base de datos
        'Si es uno se ingresa el valor del campo 'descripcion' en el campo descripcion
        'Si es cero no se ingresa valor alguno al campo descripcion.
     
    'factura, este es un consecutivo, se tiene que buscar la ultima factura y agregar 1, para tener la nueva factura.
        'Ejemplo, en base de datos la última factura es la 234, por tanto la nueva debe ser 235.
        'En la clase DataBase esta la función fnInser, si se usa esta con la propiedad m_bFolio a true te crea el nuevo registro 'padre'
        'El valor de regreso debe ser el número de la nueva factura, pero si falló la inserción, debe regresar un 0 (cero)
        
    'el nuevo credito se ingresa a la entidad 'creditos'
    
    'Ver definiciòn de tabla creditos incluida en el correo

End Function

Public Function validaDisponibilidadDeCredito(iCliente As Integer, fCreditoAdicional As Double) As Boolean
    
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    Dim fTotalOtorgado As Double
    Dim fMaximoOtorgado As Double
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOCliente"
    m_oBD.bFolio = False
    m_oBD.Parametros = iCliente
    m_oBD.fnObtenInformacion
    
    If m_oBD.cDetalle.Count > 0 Then
        
        Set cRegistro = m_oBD.cDetalle(1)
    
        Set oCampo = cRegistro(2) 'Credito otorgado actual
        fTotalOtorgado = oCampo.Valor
        
        Set oCampo = cRegistro(3) 'Maximo credito otorgado
        fMaximoOtorgado = oCampo.Valor
        
        'Si tiene creditos pendientes y el monto de credito actual + nuevo monto es <= total de credito otorgado
        If fMaximoOtorgado >= fTotalOtorgado + fCreditoAdicional Then
            'regresar true para indicar la disponiblidad de crédito
            validaDisponibilidadDeCredito = True
        Else
        'Si tiene creditos pendientes y el monto de credito actual + nuevo monto es > total de credito otorgado
            'regresar false para indicar que no hay la posiblidad de otorgar mas crédito
            validaDisponibilidadDeCredito = False
        End If
    Else
        'Si no tiene creditos pendientes o no ha tenido créditos
            'regresa true para indicar la disponiblidad de crédito
        m_oBD.Procedimiento = "CLIENTEObtenDatos"
        m_oBD.fnObtenInformacion
        Set cRegistro = m_oBD.cDetalle(1)
        Set oCampo = cRegistro(10) 'Maximo credito otorgado
        fMaximoOtorgado = oCampo.Valor
        If fMaximoOtorgado >= fCreditoAdicional Then
            validaDisponibilidadDeCredito = True
        Else
            validaDisponibilidadDeCredito = False
        End If
    End If
    Set m_oBD = Nothing
            
'COnsideraciones

    'La entidad qryvalidacredito, contiene la información del credito por cliente
    'En el campo TCredito esta el total de crédito que asta el momento un cliente tiene utilizado del total otorgado.
    'En campo maxcredito tiene el total de crédito otorgado a un cliente
    'Ver definiciòn de tabla qryvalidacredito incluida en el correo
    
End Function

Public Function validaCreditoCliente(iCliente As Integer, fCreditoAdicional As Double, fCreditoAnterior As Double) As Boolean
    
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    Dim fTotalOtorgado As Double
    Dim fMaximoOtorgado As Double
    Dim strStatus As String
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOMontoCliente"
    m_oBD.bFolio = False
    m_oBD.Parametros = iCliente
    m_oBD.fnObtenInformacion
    
    If m_oBD.cDetalle.Count > 0 Then
        
        Set cRegistro = m_oBD.cDetalle(1)
    
        Set oCampo = cRegistro(2) 'Credito otorgado actual
        fTotalOtorgado = oCampo.Valor
        
        Set oCampo = cRegistro(3) 'Maximo credito otorgado
        fMaximoOtorgado = oCampo.Valor
        
        Set oCampo = cRegistro(4) 'Status
        strStatus = oCampo.Valor
        
        If strStatus = "V" Then
            If fMaximoOtorgado >= (fTotalOtorgado + fCreditoAdicional - fCreditoAnterior) Then
                validaCreditoCliente = True
            Else
                validaCreditoCliente = False
            End If
        Else
            validaCreditoCliente = True
        End If
        
    Else
        'Si no tiene creditos pendientes o no ha tenido créditos
            'regresa true para indicar la disponiblidad de crédito
        validaCreditoCliente = True
    End If
    Set m_oBD = Nothing
            
'COnsideraciones

    'La entidad qryvalidacredito, contiene la información del credito por cliente
    'En el campo TCredito esta el total de crédito que asta el momento un cliente tiene utilizado del total otorgado.
    'En campo maxcredito tiene el total de crédito otorgado a un cliente
    'Ver definiciòn de tabla qryvalidacredito incluida en el correo
    
End Function

'Public Function datosCredito(iCliente As Integer, iFactura As Integer) As Collection
Public Function datosCredito(lFactura As Long) As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOObten"
    m_oBD.bFolio = False
    m_oBD.Parametros = lFactura
    Set m_cDatos = m_oBD.fnObtenInformacion
    datosCredito = m_oBD.bDatos
    Set m_oBD = Nothing
    
    'Este método debe buscar los datos de un crédito, dado un cliente y la factura
    'La función deja la iformación en la propiedad m_cDatos, para su consulta.
    
'Consideraciones:

    'Utilizar la clase DataBase
    'Asignar el nombre del procedimiento almacenado a la propiedad Procedimiento de la clase Database. (el sp debe tener por nombre CREDITOObtenDatos)
    'Utilizar la función fnObtenInformación de la clase DataBase para realizar la consulta.
    'Ver definiciòn de tabla creditos incluida en el correo
End Function

Public Function obtenGenerales(iFactura As Long) As Collection

    Dim cCredito As New Collection
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOObtenGenerales"
    m_oBD.bFolio = False
    m_oBD.Parametros = iFactura
    
    Set cCredito = m_oBD.fnObtenInformacion
    If cCredito.Count > 0 Then
    
        Dim cRegistro As New Collection
        Dim oCampo As New Campo
            
        Dim fCredito As Double
        Dim fAdeudo As Double
        Dim fTotalPagar As Double
        Dim fDescuento As Double
        Dim fCantidadPagada As Double
        Dim fNotaDescuento As Double
        Dim iPorcentajeFinanciamiento As Integer
        Dim iEstadoCredito As Integer
        Dim iDias As Integer
        Dim dtFecha As String
        Dim dtFechaFin As String
        Dim iElectrico As Integer
    
        Set cRegistro = cCredito(1)
        
        Set oCampo = cRegistro(5) ' Credito otorgado
        fCredito = oCampo.Valor
        
        If fCredito > 0# Then
        Set oCampo = cRegistro(6) ' Financiamiento
            iPorcentajeFinanciamiento = (oCampo.Valor / fCredito) * 100
        Else
            iPorcentajeFinanciamiento = 0
        End If
         
        Set oCampo = cRegistro(9) ' Cantidad total a pagar
        If IsNull(oCampo.Valor) Then
            fTotalPagar = 0#
        Else
            fTotalPagar = oCampo.Valor
        End If
        
        Set oCampo = cRegistro(15) ' Descuento
        If IsNull(oCampo.Valor) Then
            fDescuento = 0#
        Else
            fDescuento = oCampo.Valor
        End If
        
        Set oCampo = cRegistro(16) ' Nota Descuento
        If IsNull(oCampo.Valor) Then
            fNotaDescuento = 0#
        Else
            fNotaDescuento = oCampo.Valor
        End If
        
        Set oCampo = cRegistro(17) ' Electrico
        iElectrico = oCampo.Valor
        
        Set oCampo = cRegistro(19) ' Cantidad total pagada
        fCantidadPagada = oCampo.Valor
        
        fAdeudo = fTotalPagar - fCantidadPagada - fDescuento - fNotaDescuento
        
        Set oCampo = cRegistro(10) ' Fecha creación crédito
        dtFecha = oCampo.Valor
        
        Set oCampo = cRegistro(12) ' Fecha final crédito
        dtFechaFin = oCampo.Valor
        
        Set oCampo = cRegistro(18) ' Status
        Select Case oCampo.Valor
            
            Case Is = "V"
                iEstadoCredito = 0
                iDias = DateDiff("d", dtFecha, Format(Date, "dd/mm/yyyy")) - 1
            
            Case Is = "T"
                iEstadoCredito = 1
                'iDias = DateDiff("d", dtFechaFin, Format(Date, "dd/mm/yyyy")) - 1
                iDias = DateDiff("d", Format(dtFecha, "dd/mm/yyyy"), dtFechaFin) - 1
            Case Is = "P"
                iEstadoCredito = 2
                iDias = DateDiff("d", Format(dtFecha, "dd/mm/yyyy"), dtFechaFin) - 1
                
                'iDias = DateDiff("d", dtFechaFin, Format(Date, "dd/mm/yyyy")) - 1
            Case Is = "C"
                iEstadoCredito = 3
                iDias = DateDiff("d", dtFechaFin, Format(Date, "dd/mm/yyyy")) - 1
            Case Is = "E"
                iEstadoCredito = 4
                iDias = DateDiff("d", dtFechaFin, Format(Date, "dd/mm/yyyy")) - 1
            
        End Select
        
        cRegistro.Add oCampo.CreaCampo(adInteger, , , iPorcentajeFinanciamiento) '(20)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , iEstadoCredito)            '(21)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , iDias)                     '(22)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , fAdeudo)                   '(23)
        cRegistro.Add oCampo.CreaCampo(adInteger, , , iElectrico)                '(24)
        
    End If
    
    Set obtenGenerales = cCredito
    
    Set m_oBD = Nothing

End Function

Public Function evaluaCredito(ByRef fPagoDiario As Double, ByRef fTotalPagar As Double, ByRef fFinanciamiento As Double, ByVal fMontoCredito As Double, ByVal iPorcentajeFinanciamiento As Integer, ByVal iNumeroPagos As Integer)

    Dim strDecimales, strPagoDiario, strNumPagos, strEntero As String
    Dim iPosPunto As Integer
    Dim iMeses As Integer
    Dim iPorcentaje As Double
    'La función debe calcular el pago diario, total a pagar y el financiamiento, en función de MOnto de crédito, Porcentaje de financiamiento y No. de pagos, de al siguiente manera
    
    'El Pago Diario es igual el Monto de credito mas el porcentaje de financiamiento sobre el monto de credito entre el no. de pagos
    'If iNumeroPagos > 0 Then
'Cambio realizado el 19/07/2011
'        fPagoDiario = (fMontoCredito + (fMontoCredito * iPorcentajeFinanciamiento / 100)) / iNumeroPagos
'Cambio realizado el 19/07/2011
    'Else
    '    fPagoDiario = 0#
    'End If
'Cambio realizado el 19/07/2011
'    If iNumeroPagos < 31 Then
'        fPagoDiario = (fMontoCredito / 1000) * 44
'    Else
'        strNumPagos = CStr(iNumeroPagos / 30#)
'
'        iPosPunto = InStr(strNumPagos, ".")
'
'        If iPosPunto > 0 Then
'            strDecimales = Mid(strNumPagos, iPosPunto + 1, 2)
'
'            If Val(strDecimales) > 0 Then
'                strEntero = Mid(strNumPagos, 1, iPosPunto - 1)
'                iMeses = Val(strEntero) + 1
'            End If
'        Else
'            iMeses = Val(strNumPagos)
'        End If
'
'        fPagoDiario = (fMontoCredito + (fMontoCredito * iPorcentajeFinanciamiento / 100#)) / (26 * iMeses)
'
'    End If
    
'Cambio realizado el 19/07/2011
    fPagoDiario = (fMontoCredito + (fMontoCredito * iPorcentajeFinanciamiento / 100#)) / iNumeroPagos
    
    strPagoDiario = CStr(fPagoDiario)
    
    iPosPunto = InStr(strPagoDiario, ".")
    
    If iPosPunto > 0 Then
        strDecimales = Mid(strPagoDiario, iPosPunto + 1, 2)

        If Val(strDecimales) > 0 Then
            strEntero = Mid(strPagoDiario, 1, iPosPunto - 1)
            fPagoDiario = Val(strEntero) + 1
        End If
    End If
    
    'El total a pagar es igual al MOnto de credito mas el porcentaje de financiamiento sobre el monto de crédito.
    'fTotalPagar = fMontoCredito + (fMontoCredito * iPorcentaje / 100#)
    fTotalPagar = fMontoCredito + (fMontoCredito * iPorcentajeFinanciamiento / 100)
    'El FInanciamiento es igual al Porcentaje de finanaciamiento sobre el Monto de crédito.
    'fFinanciamiento = (fMontoCredito * iPorcentaje / 100#)
    fFinanciamiento = (fMontoCredito * iPorcentajeFinanciamiento / 100)
    
End Function

Public Function statusCatalogo() As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOEstatusCatalogo"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    statusCatalogo = m_bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function obtenPorEstatus(strEstatus As String, strFecha As String) As Boolean

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "CLIENTEGeneraCreditos"
    m_oBD.Parametros = 0 & ", '" & strEstatus & "'"
    Call m_oBD.fnInsert
    
    m_oBD.Procedimiento = "CLIENTEObtenCreditos"
    m_oBD.Parametros = 0 & ", 'V', '" & strFecha & "'"
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    obtenPorEstatus = m_bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function obtenEspectativa(strFecha As String) As Boolean

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "generaEspectativa"
    'm_oBD.Parametros = "'" & strFecha & "'"
    m_oBD.Parametros = "'01/05/2008'"
    m_oBD.bFolio = False
    m_oBD.fnInsert
    
    m_oBD.Procedimiento = "obtenEspectativa"
    m_oBD.Parametros = ""
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    obtenEspectativa = m_bDatos
    Set m_oBD = Nothing
    
End Function

Public Function elimina(iFactura As Long) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CREDITOElimina"
    m_oBD.Parametros = iFactura
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    elimina = m_bDatos
    Set m_oBD = Nothing

End Function

Public Function actualizaCreditos(strFecha As String) As Boolean

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "rpGeneraInformacionRepHandHeld"
    'm_oBD.Procedimiento = "rpGeneraInfoRepHandHeld"
    m_oBD.Parametros = "'" & strFecha & "'"
    m_oBD.bFolio = False
'MsgBox "Procede a actualizar creditos"
    m_oBD.fnInsert
    Set m_oBD = Nothing
    
End Function

Public Function corte(fEfeChe As Double, fCheque As Double, fDevo As Double, strFecha As String, iCorte As Integer, strCobradorDescansa As String, strCobradorSuple As String)

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "rpReporteGeneral"
    m_oBD.Parametros = fEfeChe & ", " & fCheque & ", " & fDevo & ", '" & strFecha & "', " & iCorte & ", '" & strCobradorDescansa & "', '" & strCobradorSuple & "'"
    m_oBD.bFolio = False
    m_oBD.fnInsert
    Set m_oBD = Nothing

End Function

Public Function obtenMoratorios()

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "obtenReglasMoratorios"
    m_oBD.Parametros = ""
    Set m_cDatos = m_oBD.fnObtenInformacion()
    m_bDatos = m_oBD.bDatos
    Set m_oBD = Nothing

End Function

Public Function actualizaMoratorios(cReglas As Collection)
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "actualizaReglasMoratorios"
    m_oBD.Parametros = ""
    m_oBD.bFolio = False
    m_oBD.cDetalle = cReglas
    m_oBD.fnInsertaDetalle
    Set m_oBD = Nothing
    
End Function

Public Function corteDiario()

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "rpCorteDiario"
    m_oBD.Parametros = "'" & Format(Date, "dd/mm/yyyy") & "'"
    m_oBD.bFolio = False
    m_oBD.fnInsert
    Set m_oBD = Nothing

End Function

Public Function obtenPagosInternet(strArchivoDatos As String) As String

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "rpEstadoCuentaInternet"
    m_oBD.Parametros = "'" & Date & "', " & "'" & Date & "'"
    m_oBD.bFolio = False
    obtenPagosInternet = m_oBD.fnObtenInformacionEnArchivo(strArchivoDatos)
    Set m_oBD = Nothing
    
End Function

Public Function obtenEstadosCuentaHH(strArchivoDatos As String) As String

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "rpHH"
    m_oBD.Parametros = "'" & Date & "', " & "'" & Date & "'"
    m_oBD.bFolio = False
    obtenEstadosCuentaHH = m_oBD.fnObtenInformacionEnArchivoUno(strArchivoDatos)
    Set m_oBD = Nothing
    
End Function


Public Function generaCreditosTemporal(Registros) As Collection

    Dim m_oBD As New DataBase
    
    m_oBD.Parametros = ""
    m_oBD.bFolio = False
    
    m_oBD.Procedimiento = "limpiaCreditosNuevosPaso"
    m_oBD.fnInsert
    
    m_oBD.Procedimiento = "insertaCreditoNuevoPaso"
    m_oBD.cDetalle = Registros
    m_oBD.fnInsertaDetalle
    
    m_oBD.Procedimiento = "obtenCreditosNuevosSolicitados"
    Set generaCreditosTemporal = m_oBD.fnObtenInformacion()
    
    Set m_oBD = Nothing

End Function
