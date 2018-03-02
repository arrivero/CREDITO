VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Empleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvariEmpleado As Integer 'local copy
Private miEstadoEmpleado As String
Private mstrNombre As String
Private mstrApPaterno As String
Private mstrApMaterno As String
Private mfNacimiento As Date
Private mstrRFC As String
Private mstrCURP As String
Private miEstadoCivil As Integer
Private miGenero As Integer
Private mstrTelefono As String
Private mstrCorreo As String
Private mdSalario As Double 'local copy

Private mstrDireccion As String
Private mstrColonia As String
Private mstrCiudad As String
Private mstrEstado As String
Private mstrCP As String

Private mstrUsuario As String
Private miPerfil As Integer
Private miEstadoUsuario As Integer

Private m_cDatos As Collection
Private m_bDatos As Boolean

Private m_cDetalle As Collection

Public Property Let iEmpleado(ByVal vData As Integer)
    mvariEmpleado = vData
End Property

Public Property Get iEmpleado() As Integer
    iEmpleado = mvariEmpleado
End Property

Public Property Let iEstadoEmpleado(ByVal vData As Integer)
    miEstadoEmpleado = vData
End Property

Public Property Get iEstadoEmpleado() As Integer
    iEstadoEmpleado = miEstadoEmpleado
End Property

Public Property Let strNombre(ByVal vData As String)
    mstrNombre = vData
End Property

Public Property Get strNombre() As String
    strNombre = mstrNombre
End Property

Public Property Let strApPaterno(ByVal vData As String)
    mstrApPaterno = vData
End Property

Public Property Get strApPaterno() As String
    strApPaterno = mstrApPaterno
End Property

Public Property Let strApMaterno(ByVal vData As String)
    mstrApMaterno = vData
End Property

Public Property Get strApMaterno() As String
    strApMaterno = mstrApMaterno
End Property

Public Property Let fNacimiento(ByVal vData As Date)
    mfNacimiento = vData
End Property

Public Property Get fNacimiento() As Date
    fNacimiento = mfNacimiento
End Property

Public Property Let strRFC(ByVal vData As String)
    mstrRFC = vData
End Property

Public Property Get strRFC() As String
    strRFC = mstrRFC
End Property

Public Property Let strCURP(ByVal vData As String)
    mstrCURP = vData
End Property

Public Property Get strCURP() As String
    strCURP = mstrCURP
End Property

Public Property Let iEstadoCivil(ByVal vData As Integer)
    miEstadoCivil = vData
End Property

Public Property Get iEstadoCivil() As Integer
    iEstadoCivil = miEstadoCivil
End Property

Public Property Let iGenero(ByVal vData As Integer)
    miGenero = vData
End Property

Public Property Get iGenero() As Integer
    iGenero = miGenero
End Property

Public Property Let strTelefono(ByVal vData As String)
    mstrTelefono = vData
End Property

Public Property Get strTelefono() As String
    strTelefono = mstrTelefono
End Property

Public Property Let strCorreo(ByVal vData As String)
    mstrCorreo = vData
End Property

Public Property Get strCorreo() As String
    strCorreo = mstrCorreo
End Property

Public Property Let dSalario(ByVal vData As Double)
    mdSalario = vData
End Property

Public Property Get dSalario() As Double
    dSalario = mdSalario
End Property

Public Property Let strDireccion(ByVal vData As String)
    mstrDireccion = vData
End Property

Public Property Get strDireccion() As String
    strDireccion = mstrDireccion
End Property

Public Property Let strColonia(ByVal vData As String)
    mstrColonia = vData
End Property

Public Property Get strColonia() As String
    strColonia = mstrColonia
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

Public Property Let strUsuario(ByVal vData As String)
    mstrUsuario = vData
End Property

Public Property Get strUsuario() As String
    strUsuario = mstrUsuario
End Property

Public Property Let iPerfil(ByVal vData As Integer)
    miPerfil = vData
End Property

Public Property Get iPerfil() As Integer
    iPerfil = miPerfil
End Property

Public Property Let iEstadoUsuario(ByVal vData As Integer)
    miEstadoUsuario = vData
End Property

Public Property Get iEstadoUsuario() As Integer
    iEstadoUsuario = miEstadoUsuario
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

Public Property Get cDetalle() As Collection
    Set cDetalle = m_cDetalle
End Property

Public Property Let cDetalle(ByVal vData As Collection)
    Set m_cDetalle = vData
End Property

Public Function catalogo() As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "obtenCatalogo"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    catalogo = m_oBD.bDatos
    Set m_oBD = Nothing

End Function

Public Function alta(cDatos As Collection) As Integer
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "altaEmpleado"
    m_oBD.Parametros = ""
    m_oBD.bFolio = True
    m_oBD.cDetalle = cDatos
    mvariEmpleado = m_oBD.fnInsert
    
    alta = mvariEmpleado
    Set m_oBD = Nothing
    
End Function

Public Function actualiza(cDatos As Collection) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "actualizaEmpleado"
    m_oBD.bFolio = False
    m_oBD.cDetalle = cDatos
    m_oBD.fnInsertaDetalle
    
    Set m_oBD = Nothing
    
End Function

Public Function Informacion(iEmpleado As Integer) As Boolean
    
    Informacion = False
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "EMPLEADOObtenDatos"
    m_oBD.Parametros = iEmpleado
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Dim oCampo As New Campo
    Dim Registro As Collection
    
    For Each Registro In m_cDatos
    
        Set oCampo = Registro(1) 'Id Empleado
        If IsNull(oCampo.Valor) Then
            mvariEmpleado = 0
        Else
            mvariEmpleado = oCampo.Valor
        End If
        
        Set oCampo = Registro(2) 'Estado del empleado
        If IsNull(oCampo.Valor) Then
            miEstadoEmpleado = 0
        Else
            miEstadoEmpleado = oCampo.Valor
        End If
        
        Set oCampo = Registro(3) 'Nombre
        If IsNull(oCampo.Valor) Then
            mstrNombre = ""
        Else
            mstrNombre = oCampo.Valor
        End If
        
        Set oCampo = Registro(4) 'AppPaterno
        If IsNull(oCampo.Valor) Then
            mstrApPaterno = ""
        Else
            mstrApPaterno = oCampo.Valor
        End If
        
        Set oCampo = Registro(5) 'AppMaterno
        If IsNull(oCampo.Valor) Then
            mstrApMaterno = ""
        Else
            mstrApMaterno = oCampo.Valor
        End If
        
        Set oCampo = Registro(6) 'Fecha Nacimiento
        If IsNull(oCampo.Valor) Then
            mfNacimiento = ""
        Else
            mfNacimiento = oCampo.Valor
        End If
        
        Set oCampo = Registro(7) 'RFC
        If IsNull(oCampo.Valor) Then
            mstrRFC = ""
        Else
            mstrRFC = oCampo.Valor
        End If
        
        Set oCampo = Registro(8) 'CURP
        If IsNull(oCampo.Valor) Then
            mstrCURP = ""
        Else
            mstrCURP = oCampo.Valor
        End If
        
        Set oCampo = Registro(9) 'eSTADO CIVIL
        If IsNull(oCampo.Valor) Then
            miEstadoCivil = 0
        Else
            miEstadoCivil = oCampo.Valor
        End If
    
        Set oCampo = Registro(10) 'gENERO
        If IsNull(oCampo.Valor) Then
            miGenero = 0
        Else
            miGenero = oCampo.Valor
        End If
    
        Set oCampo = Registro(11) 'tELEFONO
        If IsNull(oCampo.Valor) Then
            mstrTelefono = ""
        Else
            mstrTelefono = oCampo.Valor
        End If
    
        Set oCampo = Registro(12) 'CORREO
        If IsNull(oCampo.Valor) Then
            mstrCorreo = ""
        Else
            mstrCorreo = oCampo.Valor
        End If
    
        Set oCampo = Registro(13) 'sALARIO
        If IsNull(oCampo.Valor) Then
            mdSalario = 0
        Else
            mdSalario = oCampo.Valor
        End If
        
        Set oCampo = Registro(14) 'Calle Numero (Direccion)
        If IsNull(oCampo.Valor) Then
            mstrDireccion = ""
        Else
            mstrDireccion = oCampo.Valor
        End If
    
        Set oCampo = Registro(15) 'Colonia
        If IsNull(oCampo.Valor) Then
            mstrColonia = ""
        Else
            mstrColonia = oCampo.Valor
        End If
    
        Set oCampo = Registro(16) 'Ciudad
        If IsNull(oCampo.Valor) Then
            mstrCiudad = ""
        Else
            mstrCiudad = oCampo.Valor
        End If
        
        Set oCampo = Registro(17) 'Estado
        If IsNull(oCampo.Valor) Then
            mstrEstado = ""
        Else
            mstrEstado = oCampo.Valor
        End If
        
        Set oCampo = Registro(18) 'Codigo Postal
        If IsNull(oCampo.Valor) Then
            mstrCP = ""
        Else
            mstrCP = oCampo.Valor
        End If
        
        Set oCampo = Registro(19) 'Usuario
        If IsNull(oCampo.Valor) Then
            mstrUsuario = ""
        Else
            mstrUsuario = oCampo.Valor
        End If
        
        Set oCampo = Registro(20) 'Perfil
        If IsNull(oCampo.Valor) Then
            miPerfil = 0
        Else
            miPerfil = oCampo.Valor
        End If
        
        Set oCampo = Registro(21) 'Estado Usuario
        If IsNull(oCampo.Valor) Then
            miEstadoUsuario = 0
        Else
            miEstadoUsuario = oCampo.Valor
        End If
        
        Informacion = True
    
    Next Registro
    
    Set m_oBD = Nothing
      
End Function

Public Function catalogoEstado() As Boolean
    
    'Obten las entidades
    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "catalogoEstadoEmpleado"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    catalogoEstado = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function catalogoEstadoCivil() As Boolean
    
    'Obten las entidades
    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "catalogoEstadoCivil"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    catalogoEstadoCivil = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function


Public Function catalogoPerfil() As Boolean
    
    'Obten las entidades
    Dim m_oBD As New DataBase
    
    m_oBD.Procedimiento = "catalogoPerfil"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    catalogoPerfil = m_oBD.bDatos
    
    Set m_oBD = Nothing
    
End Function

Public Function obtenLista() As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "EMPLEADOObtenLista"
    m_oBD.Parametros = ""
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    obtenLista = m_oBD.bDatos
    Set m_oBD = Nothing

End Function

'Public Function informacionGeneral(iEmpleado As Integer) As Boolean
'
'    informacionGeneral = False
'
'    Dim m_oBD As New DataBase
'    m_oBD.Procedimiento = "CLIENTEObtenDatos"
'    m_oBD.Parametros = iEmpleado
'    Set m_cDatos = m_oBD.fnObtenInformacion()
'
'    Dim oCampo As New Campo
'    Dim Registro As Collection
'
'    For Each Registro In m_cDatos
'
'        Set oCampo = Registro(1) 'No. de cliente
'        If IsNull(oCampo.Valor) Then
'            mvariEmpleado = 0
'        Else
'            mvariEmpleado = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(2)
'        If IsNull(oCampo.Valor) Then
'            mstrNombre = ""
'        Else
'            mstrNombre = oCampo.Valor 'Nombre
'        End If
'
'        Set oCampo = Registro(3)
'        If IsNull(oCampo.Valor) Then
'            mstrApPaterno = ""
'        Else
'            mstrApPaterno = oCampo.Valor 'AppPaterno
'        End If
'
'        Set oCampo = Registro(4)
'        If IsNull(oCampo.Valor) Then
'            mstrDireccionCliente = ""
'        Else
'            mstrDireccionCliente = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(5)
'        If IsNull(oCampo.Valor) Then
'            mstrUbicacion = ""
'        Else
'            mstrUbicacion = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(6)
'        If IsNull(oCampo.Valor) Then
'            mstrCiudad = ""
'        Else
'            mstrCiudad = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(7)
'        If IsNull(oCampo.Valor) Then
'            mstrEstado = ""
'        Else
'            mstrEstado = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(8)
'        If IsNull(oCampo.Valor) Then
'            mstrCP = ""
'        Else
'            mstrCP = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(9)
'        If IsNull(oCampo.Valor) Then
'            mstrTelefonoCliente = ""
'        Else
'            mstrTelefonoCliente = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(10)
'        If IsNull(oCampo.Valor) Then
'            mdCreditoMaximo = 0#
'        Else
'            mdCreditoMaximo = oCampo.Valor
'        End If
'
'        Set oCampo = Registro(11)
'        If IsNull(oCampo.Valor) Then
'            miAtrasosPermitidos = 0
'        Else
'            miAtrasosPermitidos = oCampo.Valor
'        End If
'
'        informacionGeneral = True
'
'    Next Registro
'
'    Set m_oBD = Nothing
'
'End Function

'Public Function listaNameApellido(strLikeNombre As String, strApellido As String) As Boolean
'
'    Dim m_oBD As New DataBase
'    m_oBD.Procedimiento = "CLIENTEListaNameApellido"
'    m_oBD.Parametros = "'" & strLikeNombre & "', '" & strApellido & "'"
'
'    Set m_cDatos = m_oBD.fnObtenInformacion()
'    listaNameApellido = m_oBD.bDatos
'    Set m_oBD = Nothing
'
'End Function
'
'Public Function clienteBusca(iSalon As Integer, strCliente As String) As Boolean
'
'    Dim m_oBD As New DataBase
'    m_oBD.Procedimiento = "clienteObten"
'    m_oBD.Parametros = iSalon & ", '" & strCliente & "'"
'
'    Set m_cDatos = m_oBD.fnObtenInformacion()
'    clienteBusca = m_oBD.bDatos
'    Set m_oBD = Nothing
'
'End Function

