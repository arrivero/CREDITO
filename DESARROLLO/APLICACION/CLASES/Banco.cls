VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private m_oBD As DataBase
Private m_bDatos As Boolean
Private m_cDatos As Collection

Public Property Get cDatos() As Collection
    Set cDatos = m_cDatos
End Property

Public Property Let cDatos(ByRef vData As Collection)
    Set m_cDatos = vData
End Property

'Public Property Let oBD(ByVal vData As DataBase)
'    Set m_oBD = vData
'End Property

Public Function fnCatalogo() As Boolean
    
    'Obten las entidades
    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "catalogoBancos"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    fnCatalogo = m_oBD.bDatos
    
    Set m_oBD = Nothing
End Function

Public Function buscaConNombre(strBanco As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "buscaNombreBanco"
    m_oBD.Parametros = "'" & strBanco & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    buscaConNombre = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function creaNuevo(strBanco As String) As Integer
    
    Dim cCuentaBanco As New Collection
    Dim Registro As New Collection
    Dim oCampo As New Campo
    Dim m_oBD As New DataBase

    'Registro.Add oCampo.CreaCampo(adInteger, , , 0)
    Registro.Add oCampo.CreaCampo(adInteger, , , strBanco)
    
    cCuentaBanco.Add Registro
    
    m_oBD.Procedimiento = "altaBanco"
    m_oBD.Parametros = 0 & ", '" & strBanco & "'"
    m_oBD.bFolio = True
    m_oBD.cDetalle = cCuentaBanco
    creaNuevo = m_oBD.fnInsert
    Set m_oBD = Nothing

End Function
