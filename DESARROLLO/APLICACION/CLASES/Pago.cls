VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public Property Get bDatos() As Boolean
    bDatos = m_bDatos
End Property

Public Property Let bDatos(ByVal vData As Boolean)
    m_bDatos = vData
End Property

Public Function obtenPagos(strUsuario As String, strFecha As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PAGOObten"
    m_oBD.Parametros = "'" & strUsuario & "', '" & strFecha & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    obtenPagos = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function obtenPagosTmp() As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PAGO_TEMPObten"
    m_oBD.Parametros = "alex"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    obtenPagosTmp = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function obtenCreditoPagos(iFolio As Long) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PAGOObtenPagosCredito"
    m_oBD.Parametros = iFolio
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    obtenCreditoPagos = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function grabaPagos(cPagos As Collection) As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.bFolio = False
    m_oBD.Procedimiento = "limpiaPagosTmp"
    m_oBD.Parametros = ""
    m_oBD.fnInsert
    
    m_oBD.Procedimiento = "PAGOGraba"
    m_oBD.cDetalle = cPagos
    m_oBD.fnInsertaDetalle
    grabaPagos = True
    
    Set m_oBD = Nothing
    
End Function

Public Function grabaPagosTmp(cPagos As Collection) As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.bFolio = False
    m_oBD.Procedimiento = "PAGO_TEMPLimpiaPagos"
    m_oBD.Parametros = ""
    m_oBD.fnInsert
    
    m_oBD.Procedimiento = "preparaPago"
    m_oBD.cDetalle = cPagos
    m_oBD.fnInsertaDetalle
    grabaPagosTmp = True
    
    Set m_oBD = Nothing
    
End Function

Public Function registraPagos() As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.bFolio = False
    m_oBD.Procedimiento = "registraPagos"
    m_oBD.Parametros = ""
    m_oBD.fnInsert
    
    registraPagos = True
    
    Set m_oBD = Nothing
    
End Function

Public Function actualizaPagos(cPagos As Collection) As Boolean
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PAGOActualiza"
    m_oBD.bFolio = False
    m_oBD.cDetalle = cPagos
    m_oBD.fnInsertaDetalle
    actualizaPagos = True
    Set m_oBD = Nothing
    
End Function

Public Function elimintaPagosFolio(iFolio As Long, iCliente As Integer)

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PAGOFolioElimina"
    m_oBD.bFolio = False
    m_oBD.Parametros = iFolio & ", " & iCliente
    m_oBD.fnInsert
    elimintaPagosFolio = True
    Set m_oBD = Nothing

End Function

Public Function registrado(lFolio As Long, strFecha As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "PAGOEstaRegistrado"
    m_oBD.bFolio = False
    m_oBD.Parametros = lFolio & ", '" & strFecha & "'"
    m_oBD.fnObtenInformacion
    registrado = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function registraSobrante(cSobrantes As Collection, strFecha As String) As Boolean

    Dim m_oBD As New DataBase
    
    'Elimina los sobrantes que haya de esta fecha
    m_oBD.Procedimiento = "SOBRANTESElimina"
    m_oBD.Parametros = "'" & strFecha & "'"
    m_oBD.bFolio = False
    m_oBD.fnInsert
    
    m_oBD.Procedimiento = "SOBRANTESRegistra"
    m_oBD.cDetalle = cSobrantes
    m_oBD.fnInsertaDetalle
    registraSobrante = True
    Set m_oBD = Nothing

End Function

Public Function obtenSobrante(strFecha As String) As Double

    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "SOBRANTESObten"
    m_oBD.bFolio = False
    m_oBD.Parametros = "'" & strFecha & "'"
    Set m_cDatos = m_oBD.fnObtenInformacion

    Set cRegistro = m_cDatos(1)
    Set oCampo = cRegistro(1)
    
    obtenSobrante = oCampo.Valor
    
    Set m_oBD = Nothing
    
End Function

Public Function obtenResumenDiarioMontos(ByRef dEfectivo As Double, ByRef dCheques As Double, ByRef dDevolucion As Double, strFecha As String) As Double

    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "obtenResumenDiarioMontos"
    m_oBD.bFolio = False
    m_oBD.Parametros = "'" & strFecha & "'"
    Set m_cDatos = m_oBD.fnObtenInformacion

    If m_cDatos.Count > 0 Then
    
        Set cRegistro = m_cDatos(1)
        Set oCampo = cRegistro(1)
        dEfectivo = oCampo.Valor
        
        Set oCampo = cRegistro(2)
        dCheques = oCampo.Valor
        
        Set oCampo = cRegistro(3)
        dDevolucion = oCampo.Valor
        
    Else
    
        dEfectivo = 0#
        dCheques = 0#
        dDevolucion = 0#
        
    End If
    
    Set m_oBD = Nothing
    
End Function

Public Function obtenSobrantes(strFecha As String) As Collection

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "ObtenSOBRANTES"
    m_oBD.bFolio = False
    m_oBD.Parametros = "'" & strFecha & "'"
    Set obtenSobrantes = m_oBD.fnObtenInformacion

    Set m_oBD = Nothing
    
End Function

Public Function acumuladoPorcentajes(strFechaInicial As String, strFechaFinal As String, iPeriodo As Integer, iMontoNo As Integer, iFuncion As Integer) As Collection

    Dim m_oBD As New DataBase

    If iFuncion = 1 Then
        m_oBD.Procedimiento = "rpEstadisticoPorcentaje"
        m_oBD.Parametros = "'" & strFechaInicial & "', '" & strFechaFinal & "', " & iPeriodo & ", " & iMontoNo
        m_oBD.bFolio = False
        m_oBD.fnInsert

        m_oBD.Procedimiento = "rpObtenEstadisticoPorcentaje " & iPeriodo
        m_oBD.Parametros = ""
        Set m_cDatos = m_oBD.fnObtenInformacion()
    Else

        m_oBD.Procedimiento = "rpEstadisticoAcumuladoAnual"
        m_oBD.Parametros = ""
        Set m_cDatos = m_oBD.fnObtenInformacion()

    End If

    Set acumuladoPorcentajes = m_cDatos

End Function

Public Function acumuladoAnual() As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "rpEstadisticoAcumuladoAnual"
    m_oBD.Parametros = ""
    Set m_cDatos = m_oBD.fnObtenInformacion()

    Set acumuladoAnual = m_cDatos
    
    Set m_oBD = Nothing

End Function

Public Function crecimientoAnual() As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "crecimientoAnual"
    m_oBD.Parametros = ""
    m_oBD.bFolio = False
    m_oBD.fnInsert
    

    m_oBD.Procedimiento = "crecimientoAnualObten"
    Set m_cDatos = m_oBD.fnObtenInformacion()

    Set crecimientoAnual = m_cDatos
    
    Set m_oBD = Nothing

End Function

Public Function crecimientoTotal() As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "crecimientoTotal"
    m_oBD.Parametros = ""
    Set m_cDatos = m_oBD.fnObtenInformacion()

    Set crecimientoTotal = m_cDatos
    
    Set m_oBD = Nothing

End Function


Public Function comparativoPrestamoCobranza() As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "comparativoPrestamoCobranza"
    m_oBD.Parametros = ""
    m_oBD.bFolio = False
    m_oBD.fnInsert
    

    m_oBD.Procedimiento = "comparativoPrestamoCobranzaObten"
    Set m_cDatos = m_oBD.fnObtenInformacion()

    Set comparativoPrestamoCobranza = m_cDatos
    
    Set m_oBD = Nothing

End Function

Public Function analisisCredito(strCobrador As String) As Collection

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "rpAnalisisCredito"
    m_oBD.Parametros = "'" & strCobrador & "'"
    m_oBD.bFolio = False
    m_oBD.fnInsert
    

    m_oBD.Procedimiento = "rpAnalisisGreditoObten"
    m_oBD.Parametros = ""
    Set m_cDatos = m_oBD.fnObtenInformacion()

    Set analisisCredito = m_cDatos
    
    Set m_oBD = Nothing

End Function

Public Function asignaRuta(strFecha As String, strCobradorDescansa As String, strCobradorSuple As String)

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "asignaRuta"
    m_oBD.bFolio = False
    m_oBD.Parametros = "'" & strFecha & "', '" & strCobradorDescansa & "', '" & strCobradorSuple & "'"
    m_oBD.fnInsert
    
    Set m_oBD = Nothing
    
End Function
