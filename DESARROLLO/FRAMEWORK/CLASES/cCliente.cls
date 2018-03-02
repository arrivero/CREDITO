VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCliente"
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
Public mfNacimiento As Date
Public mstrRFC As String
Public msiEstadoCliente As String

Public mstrDireccionCliente As String
Public mstrTelefonoCliente As String
Public mintSalon As Integer

Public moDireccion As New cDireccion
Public moTelefono As New cTelefono
    
Const CLIENTE = 1
Const Nombre = 2
Const APPATERNO = 3
Const APMATERNO = 4
Const FECHA_NACIMIENTO = 5
Const RFC = 6
Const ESTADO_CLIENTE = 7
Const DIRECCION = 5
Const TELEFONO = 6
Private mvardSaldo As Double 'local copy

'Private m_oBD As DataBase
Private m_bDatos As Boolean
Private m_cDatos As Collection
Private m_cDetalle As Collection
Private m_cDireccion As Collection
Private m_cTelefono As Collection
Private m_cCorreo As Collection

Public Property Get cDetalle() As Collection
    Set cDetalle = m_cDetalle
End Property

Public Property Let cDetalle(ByVal vData As Collection)
    Set m_cDetalle = vData
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

'Public Property Let oBD(ByVal vData As DataBase)
'    Set m_oBD = vData
'End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
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

Public Property Let strDireccion(ByVal vData As String)
    mstrDireccionCliente = vData
End Property

Public Property Get strDireccion() As String
    strDireccion = mstrDireccionCliente
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

Public Function clienteBusca(iSalon As Integer, strCliente As String) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "clienteObten"
    m_oBD.Parametros = iSalon & ", '" & strCliente & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    clienteBusca = m_oBD.bDatos
    Set m_oBD = Nothing

End Function

Public Function altaCliente() As Integer
    
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "altaCliente"
    m_oBD.Parametros = ""
    m_oBD.bFolio = True
    m_oBD.cDetalle = m_cDetalle
    mvariCliente = m_oBD.fnInsert
    
    altaCliente = mvariCliente
    Set m_oBD = Nothing
    
End Function

Public Function actualizaCliente(cDirecciones As Collection, _
                                 cTelefonos As Collection, _
                                 cCorreos As Collection) As Boolean

    Dim m_oBD As New DataBase
    'Inserta la cita
    m_oBD.Procedimiento = "actualizaCliente"
    m_oBD.Parametros = "'" & mstrNombre & "', '" & mstrApPaterno & "', '" & mstrApMaterno & "', '" & mfNacimiento & "', '" & mstrRFC & "', " & msiEstadoCliente & ", " & mvariCliente
        
    m_oBD.bFolio = False
    m_oBD.cDetalle = m_cDetalle
    'm_oBD.fnInsert
    m_oBD.fnInsertaDetalle
    
    'Elimina Direcciones
    m_oBD.Procedimiento = "eliminaDireccionCliente"
    m_oBD.Parametros = mvariCliente
    m_oBD.bFolio = False
    m_oBD.fnInsert
    
    'Elimina Telefonos
    m_oBD.Procedimiento = "eliminaTelefonoCliente"
    m_oBD.Parametros = mvariCliente
    m_oBD.fnInsert
    
    'Elimina correos
    m_oBD.Procedimiento = "eliminaCorreoCliente"
    m_oBD.Parametros = mvariCliente
    m_oBD.fnInsert
    
    Call registraDatos(cDirecciones, cTelefonos, cCorreos)
    Set m_oBD = Nothing
    
End Function

Public Function registraDatos(cDirecciones As Collection, _
                              cTelefonos As Collection, _
                              cCorreos As Collection)
    
    Dim m_oBD As New DataBase
    'Registra las direcciones
    m_oBD.Procedimiento = "registraDireccionCliente"
    m_oBD.Parametros = ""
    m_oBD.cDetalle = cDirecciones
    m_oBD.fnInsertaDetalle

    'Registra los telefonos
    m_oBD.Procedimiento = "registraTelefonoCliente"
    m_oBD.cDetalle = cTelefonos
    m_oBD.fnInsertaDetalle
    
    'Registra los CORREO
    m_oBD.Procedimiento = "registraCorreoCliente"
    m_oBD.cDetalle = cCorreos
    m_oBD.fnInsertaDetalle
    Set m_oBD = Nothing
        
End Function
                                    
Public Function fnInformacion(iSalon As Integer, iCliente As Integer, iCual As Integer)
   
    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "clienteIdObten"
    m_oBD.Parametros = iSalon & ", " & iCliente & ", " & iCual
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    Dim oCampo As New Campo
    Dim Registro As Collection
    
    For Each Registro In m_cDatos
    
        Set oCampo = Registro(1)
        If IsNull(oCampo.Valor) Then
            mstrNombre = ""
        Else
            mstrNombre = oCampo.Valor 'Nombre
        End If
        
        Set oCampo = Registro(2)
        If IsNull(oCampo.Valor) Then
            mstrApPaterno = ""
        Else
            mstrApPaterno = oCampo.Valor 'AppPaterno
        End If
        
        Set oCampo = Registro(3)
        If IsNull(oCampo.Valor) Then
            mstrApMaterno = ""
        Else
            mstrApMaterno = oCampo.Valor 'AppMaterno
        End If
        
        Set oCampo = Registro(4)
        If IsNull(oCampo.Valor) Then
            mfNacimiento = Date
        Else
            mfNacimiento = oCampo.Valor
        End If
        
        Set oCampo = Registro(5)
        If IsNull(oCampo.Valor) Then
            mstrRFC = ""
        Else
            mstrRFC = oCampo.Valor
        End If
        
        Set oCampo = Registro(6)
        If IsNull(oCampo.Valor) Then
            msiEstadoCliente = ""
        Else
            msiEstadoCliente = oCampo.Valor
        End If
        
        Set oCampo = Registro(7)
        If IsNull(oCampo.Valor) Then
            mvardSaldo = 0
        Else
            mvardSaldo = oCampo.Valor
        End If
    
    Next Registro
    Set m_oBD = Nothing
      
End Function

Public Function direcciones(iCliente As Integer) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "DireccionesCliente"
    m_oBD.Parametros = iCliente
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    direcciones = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function telefonos(iCliente As Integer) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "TelefonosCliente"
    m_oBD.Parametros = iCliente
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    telefonos = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function correos(iCliente As Integer) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "CorreosCliente"
    m_oBD.Parametros = iCliente
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    correos = m_oBD.bDatos
    Set m_oBD = Nothing
    
End Function

Public Function comparativoClientes(iSalon As Integer, _
                                    strFechaInicial As String, _
                                    strFechaFinal As String, _
                                    ByRef arrValores() As Variant, _
                                    bServicios As Boolean) As Boolean

    Dim m_oBD As New DataBase
    m_oBD.Procedimiento = "ComparativoIngresosClientes"
    m_oBD.Parametros = iSalon & ", '" & strFechaInicial & "', '" & strFechaFinal & "'"
    
    Set m_cDatos = m_oBD.fnObtenInformacion()
    
    
    If m_oBD.bDatos = True Then
    
        Dim iRenglon As Integer
        Dim oCampo As New Campo
        Dim Registro As Collection
        ReDim arrValores(1 To 11 + 1, 1 To 2)
        'ReDim arrValores(1 To m_cDatos.Count + 1, 1 To 2)
        
        arrValores(1, 1) = "" 'Servicio
        arrValores(1, 2) = "Ingreso" 'Ingreso - monto
        
        iRenglon = 2
        For Each Registro In m_cDatos
            
            Set oCampo = Registro(1)
            arrValores(iRenglon, 1) = oCampo.Valor 'Servicio
            
            Set oCampo = Registro(2)
            arrValores(iRenglon, 2) = oCampo.Valor 'Ingreso - monto
        
            iRenglon = iRenglon + 1
            
            If iRenglon = 12 Then
                Exit For
            End If
            
        Next Registro
        
        comparativoClientes = True
    
    Else
        comparativoClientes = False
    End If
    
    Set m_oBD = Nothing

End Function

