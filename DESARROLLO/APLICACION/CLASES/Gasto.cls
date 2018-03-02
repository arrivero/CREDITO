VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Gasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bDatos As Boolean
Private m_cDatos As Collection
Private m_cDetalle As Collection

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

Public Function registraGasto(cGastoDatos As Collection) As Integer

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "GASTORepLimpia"
    m_oBD.Parametros = ""
    m_oBD.bFolio = False
    Call m_oBD.fnInsert
    
    m_oBD.Procedimiento = "GASTOElimina"
    m_oBD.Parametros = "'" & Date & "'"
    m_oBD.bFolio = False
    Call m_oBD.fnInsert
        
    m_oBD.Procedimiento = "GASTORegistra"
    m_oBD.cDetalle = cGastoDatos
    Call m_oBD.fnInsertaDetalle
    
    Set m_oBD = Nothing

End Function

Public Function deposito(strFecha As String, fImporte As Double)

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "DEPOSITORegistra"
    m_oBD.Parametros = "'" & strFecha & "', " & fImporte
    m_oBD.bFolio = False
    Call m_oBD.fnInsert
    
    Set m_oBD = Nothing

End Function

Public Function actualizaGasto(strFecha As String, cGastoDatos As Collection) As Integer

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "GASTOElimina"
    m_oBD.Parametros = "'" & strFecha & "'"
    m_oBD.bFolio = False
    Call m_oBD.fnInsert
    
    m_oBD.Procedimiento = "GASTORegistra"
    m_oBD.cDetalle = cGastoDatos
    Call m_oBD.fnInsertaDetalle
    
    Set m_oBD = Nothing

End Function

Public Function actualizaGastoInterno(strFecha As String, cGastoDatos As Collection) As Integer

    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "GASTOInternoElimina"
    m_oBD.Parametros = "'" & strFecha & "'"
    m_oBD.bFolio = False
    Call m_oBD.fnInsert
    
    m_oBD.Procedimiento = "GASTORegistra"
    m_oBD.cDetalle = cGastoDatos
    Call m_oBD.fnInsertaDetalle
    
    Set m_oBD = Nothing

End Function

Public Function obtenGastos(strFecha As String, iTrabajoInterno As Integer) As Boolean

    Dim oBD As New DataBase
    oBD.Procedimiento = "GASTOObten"
    oBD.Parametros = "'" & strFecha & "', " & iTrabajoInterno
    Set m_cDatos = oBD.fnObtenInformacion()
    obtenGastos = oBD.bDatos
    Set oBD = Nothing

End Function

Public Function catalogo() As Boolean
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "GASTOCatalogo"
    Set m_cDatos = oBD.fnObtenInformacion()
    catalogo = oBD.bDatos
    Set oBD = Nothing
    
End Function

'Public Function actualizaCredito(cCreditoDatos As Collection) As Integer
'
'    Dim m_oBD As New DataBase
'    m_oBD.Procedimiento = "CREDITOActualiza"
'    m_oBD.bFolio = False
'    m_oBD.cDetalle = m_cDetalle
'    m_oBD.fnInsertaDetalle
'    Set m_oBD = Nothing
'
'End Function




