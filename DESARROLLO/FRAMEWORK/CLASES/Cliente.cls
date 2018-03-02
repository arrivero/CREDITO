VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'OBJETOS PARA ACCESO A LA BASE DE DATOS

Private mvariCliente As Integer 'local copy

Private mvarmvardDia As Date 'local copy

Public mstrNombre As String
Public mstrApPaterno As String
Public mstrApMaterno As String
Public mstrUbicacion As String
Public mstrCiudad As String
Public mstrEntidadFederativa As String
Public mstrCP As String
Public mdCreditoMaximo As Double
Public miAtrasosPermitidos As Integer
Public mstrColonia As String
Public mstrEstado As String


Public mfNacimiento As Date
Public mstrRFC As String
Public msiEstadoCliente As String

Public mstrDireccionCliente As String
Public mstrTelefonoCliente As String
Public mintSalon As Integer

Const Cliente = 1
Const Nombre = 2
Const APPATERNO = 3
Const APMATERNO = 4
Const FECHA_NACIMIENTO = 5
Const RFC = 6
Const ESTADO_CLIENTE = 7
Const DIRECCION = 5
Const TELEFONO = 6
Private mvardSaldo As Double 'local copy

Private m_bDatos As Boolean
Private m_cDatos As Collection
Private m_cDetalle As Collection
Private m_cDireccion As Collection
Private m_cTelefono As Collection
Private m_cCorreo As Collection

Private m_cCreditos As Collection

Public Property Get cDetalle() As Collection
    Set cDetalle = m_cDetalle
End Property

Public Property Let cDetalle(ByVal vData As Collection)
    Set m_cDetalle = vData
End Property

Public Property Get cCreditos() As Collection
    Set cCreditos = m_cCreditos
End Property

Public Property Let cCreditos(ByVal vData As Collection)
    Set m_cCreditos = vData
End Property

Public Property Get cDireccion() As Collection
    Set cDireccion = m_cDireccion
End Property

Public Property Let cDireccion(ByVal vData As Collection)
    Set m_cDireccion = vData
End Property

Public Property Get cTelefono() As Collection
    Set cTelefono = m_cTelefono
End Property

Public Property Let cTelefono(ByVal vData As Collection)
    Set m_cTelefono = vData
End Property

Public Property Get cCorreo() As Collection
    Set cCorreo = m_cCorreo
End Property

Public Property Let cCorreo(ByVal vData As Collection)
    Set m_cCorreo = vData
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

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let dCreditoMaximo(ByVal vData As Double)
    mdCreditoMaximo = vData
End Property

Public Property Get dCreditoMaximo() As Double
    dCreditoMaximo = mdCreditoMaximo
End Property

Public Property Let dFechaServicio(ByVal vData As Date)
    mvarmvardDia = vData
End Property

Public Property Get dFechaServicio() As Date
    dFechaServicio = mvarmvardDia
End Property

Public Property Let iCliente(ByVal vData As Integer)
    mvariCliente = vData
End Property

Public Property Get iCliente() As Integer
    iCliente = mvariCliente
End Property

Public Property Let iAtrasosPermitidos(ByVal vData As Integer)
    miAtrasosPermitidos = vData
End Property

Public Property Get iAtrasosPermitidos() As Integer
    iAtrasosPermitidos = miAtrasosPermitidos
End Property

Public Property Let strDireccion(ByVal vData As String)
    mstrDireccionCliente = vData
End Property

Public Property Get strDireccion() As String
    strDireccion = mstrDireccionCliente
End Property


Public Property Let strUbicacion(ByVal vData As String)
    mstrUbicacion = vData
End Property

Public Property Get strUbicacion() As String
    strUbicacion = mstrUbicacion
End Property

Public Property Let strCiudad(ByVal vData As String)
    mstrCiudad = vData
End Property

Public Property Get strCiudad() As String
    strCiudad = mstrCiudad
End Property

Public Property Let strEstado(ByVal vData As String)
    mstrEstado = vData
End Property

Public Property Get strEstado() As String
    strEstado = mstrEstado
End Property
Public Property Let strCP(ByVal vData As String)
    mstrCP = vData
End Property

Public Property Get strCP() As String
    strCP = mstrCP
End Property

Public Property Let strTelefono(ByVal vData As String)
    mstrTelefonoCliente = vData
End Property

Public Property Get strTelefono() As String
    strTelefono = mstrTelefonoCliente
End Property

Public Property Let iSalon(ByVal vData As Integer)
    mintSalon = vData
End Property

Public Property Get iSalon() As Integer
    iSalon = mintSalon
End Property

Public Function catalogo() As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CatalogoCliente"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    catalogo = m_oBD.bDatos
    Set m_oBD = Nothing

End Function

Public Function listaLimiteCredito() As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEListaLimiteCredito"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    listaLimiteCredito = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function listaLikeName(strLikeNombre As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEListaLikeName"
    m_oBD.Parametros = "'" & strLikeNombre & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    listaLikeName = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function listaNameApellido(strLikeNombre As String, strApellido As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEListaNameApellido"
    m_oBD.Parametros = "'" & strLikeNombre & "', '" & strApellido & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    listaNameApellido = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function clienteBusca(iSalon As Integer, strCliente As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "clienteObten"
    m_oBD.Parametros = iSalon & ", '" & strCliente & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    clienteBusca = m_oBD.bDatos
    Set m_oBD = Nothing

End Function

Public Function altaCliente(cDatosCliente As Collection) As Integer
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEAlta"
    m_oBD.Parametros = ""
    m_oBD.bFolio = True
    m_oBD.cDetalle = cDatosCliente
    mvariCliente = m_oBD.fnInsert
    
    altaCliente = mvariCliente
    Set m_oBD = Nothing
    
End Function

Public Function actualizaCliente() As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEActualiza"
    m_oBD.bFolio = False
    m_oBD.cDetalle = m_cDetalle
    m_oBD.fnInsertaDetalle
    
    Set m_oBD = Nothing
    
End Function

Public Function fnInformacion(iCliente As Integer, strFecha As String) As Boolean
    
    'Dim cCreditos As New Collection
    'Dim cCredito As New Collection
    
    fnInformacion = False
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEObtenDatos"
    m_oBD.Parametros = iCliente
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Dim oCampo As New Campo
    Dim Registro As Collection
    
    For Each Registro In m_cDatos
    
        Set oCampo = Registro(1) 'No. de cliente
        If IsNull(oCampo.Valor) Then
            mvariCliente = 0
        Else
            mvariCliente = oCampo.Valor
        End If
        
        Set oCampo = Registro(2)
        If IsNull(oCampo.Valor) Then
            mstrNombre = ""
        Else
            mstrNombre = oCampo.Valor 'Nombre
        End If
        
        Set oCampo = Registro(3)
        If IsNull(oCampo.Valor) Then
            mstrApPaterno = ""
        Else
            mstrApPaterno = oCampo.Valor 'AppPaterno
        End If
        
        Set oCampo = Registro(4)
        If IsNull(oCampo.Valor) Then
            mstrDireccionCliente = ""
        Else
            mstrDireccionCliente = oCampo.Valor
        End If
        
        Set oCampo = Registro(5)
        If IsNull(oCampo.Valor) Then
            mstrUbicacion = ""
        Else
            mstrUbicacion = oCampo.Valor
        End If
        
        Set oCampo = Registro(6)
        If IsNull(oCampo.Valor) Then
            mstrCiudad = ""
        Else
            mstrCiudad = oCampo.Valor
        End If
        
        Set oCampo = Registro(7)
        If IsNull(oCampo.Valor) Then
            mstrEstado = ""
        Else
            mstrEstado = oCampo.Valor
        End If
    
        Set oCampo = Registro(8)
        If IsNull(oCampo.Valor) Then
            mstrCP = ""
        Else
            mstrCP = oCampo.Valor
        End If
    
        Set oCampo = Registro(9)
        If IsNull(oCampo.Valor) Then
            mstrTelefonoCliente = ""
        Else
            mstrTelefonoCliente = oCampo.Valor
        End If
    
        Set oCampo = Registro(10)
        If IsNull(oCampo.Valor) Then
            mdCreditoMaximo = 0#
        Else
            mdCreditoMaximo = oCampo.Valor
        End If
    
        Set oCampo = Registro(11)
        If IsNull(oCampo.Valor) Then
            miAtrasosPermitidos = 0
        Else
            miAtrasosPermitidos = oCampo.Valor
        End If
    
        Set oCampo = Registro(12)
        If IsNull(oCampo.Valor) Then
            mstrApMaterno = ""
        Else
            mstrApMaterno = oCampo.Valor
        End If
    
        Set oCampo = Registro(13)
        If IsNull(oCampo.Valor) Then
            mstrColonia = ""
        Else
            mstrColonia = oCampo.Valor
        End If
    
        Set oCampo = Registro(14)
        If IsNull(oCampo.Valor) Then
            mstrEntidadFederativa = ""
        Else
            mstrEntidadFederativa = oCampo.Valor
        End If
    
        fnInformacion = True
    
    Next Registro
    
    m_oBD.Procedimiento = "CLIENTEGeneraCreditos"
    m_oBD.Parametros = iCliente & ", 'V'"
    Call m_oBD.fnInsert
    
    Set m_oBD = Nothing
      
End Function

Public Function creditosPorEstatus(iCliente As Integer, strFecha As String, strStatus As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEObtenCreditos"
    m_oBD.Parametros = iCliente & ", '" & strStatus & "', '" & strFecha & "'"
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    If m_cDatos.Count > 0 Then
        creditosPorEstatus = True
    Else
        creditosPorEstatus = False
    End If
    
    Set m_oBD = Nothing
    
End Function

Public Function informacionGeneral(iCliente As Integer) As Boolean
    
    informacionGeneral = False
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CLIENTEObtenDatos"
    m_oBD.Parametros = iCliente
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Dim oCampo As New Campo
    Dim Registro As Collection
    
    For Each Registro In m_cDatos
    
        Set oCampo = Registro(1) 'No. de cliente
        If IsNull(oCampo.Valor) Then
            mvariCliente = 0
        Else
            mvariCliente = oCampo.Valor
        End If
        
        Set oCampo = Registro(2)
        If IsNull(oCampo.Valor) Then
            mstrNombre = ""
        Else
            mstrNombre = oCampo.Valor 'Nombre
        End If
        
        Set oCampo = Registro(3)
        If IsNull(oCampo.Valor) Then
            mstrApPaterno = ""
        Else
            mstrApPaterno = oCampo.Valor 'AppPaterno
        End If
        
        Set oCampo = Registro(4)
        If IsNull(oCampo.Valor) Then
            mstrDireccionCliente = ""
        Else
            mstrDireccionCliente = oCampo.Valor
        End If
        
        Set oCampo = Registro(5)
        If IsNull(oCampo.Valor) Then
            mstrUbicacion = ""
        Else
            mstrUbicacion = oCampo.Valor
        End If
        
        Set oCampo = Registro(6)
        If IsNull(oCampo.Valor) Then
            mstrCiudad = ""
        Else
            mstrCiudad = oCampo.Valor
        End If
        
        Set oCampo = Registro(7)
        If IsNull(oCampo.Valor) Then
            mstrEstado = ""
        Else
            mstrEstado = oCampo.Valor
        End If
    
        Set oCampo = Registro(8)
        If IsNull(oCampo.Valor) Then
            mstrCP = ""
        Else
            mstrCP = oCampo.Valor
        End If
    
        Set oCampo = Registro(9)
        If IsNull(oCampo.Valor) Then
            mstrTelefonoCliente = ""
        Else
            mstrTelefonoCliente = oCampo.Valor
        End If
    
        Set oCampo = Registro(10)
        If IsNull(oCampo.Valor) Then
            mdCreditoMaximo = 0#
        Else
            mdCreditoMaximo = oCampo.Valor
        End If
    
        Set oCampo = Registro(11)
        If IsNull(oCampo.Valor) Then
            miAtrasosPermitidos = 0
        Else
            miAtrasosPermitidos = oCampo.Valor
        End If
    
        informacionGeneral = True
    
    Next Registro
    
    Set m_oBD = Nothing
      
End Function

