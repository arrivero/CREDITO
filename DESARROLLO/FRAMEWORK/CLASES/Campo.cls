VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Campo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'variables locales para almacenar los valores de las propiedades
Private m_iTipoDato As Integer
Private m_iLongitud As Integer
Private m_iPosicionesDecimal As Integer
Private m_vValor As Variant

Public Property Let Valor(ByVal vData As Variant)
    m_vValor = vData
End Property

Public Property Get Valor() As Variant
    Valor = m_vValor
End Property

Public Function CreaCampo(Optional ByVal TipoDato As Integer = "1", Optional ByVal Longitud As Integer = "0", Optional ByVal PosicionesDecimal As Integer = "0", Optional ByVal vvalor As Variant) As Campo
    
    Dim oCampo As New Campo
    
    oCampo.iTipoDato = TipoDato
    oCampo.iLongitud = Longitud
    oCampo.iPosicionesDecimal = PosicionesDecimal
    oCampo.Valor = vvalor
    
    Set CreaCampo = oCampo
    
End Function

Public Property Let iPosicionesDecimal(ByVal vData As Integer)
    m_iPosicionesDecimal = vData
End Property

Public Property Get iPosicionesDecimal() As Integer
    iPosicionesDecimal = m_iPosicionesDecimal
End Property

Public Property Let iLongitud(ByVal vData As Integer)
    m_iLongitud = vData
End Property

Public Property Get iLongitud() As Integer
    iLongitud = m_iLongitud
End Property

Public Property Let iTipoDato(ByVal vData As Integer)
    m_iTipoDato = vData
End Property

Public Property Get iTipoDato() As Integer
    iTipoDato = m_iTipoDato
End Property



