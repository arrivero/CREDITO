VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E7390F601D3"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'local variable(s)

Private mvarstrNombre As String
Private mvarstrAppPaterno As String
Private mvarstrAppMaterno As String
Private mvariId As Integer
Private mvarstrId As String
Private mvarstrPsw As String
Private mvariTipo As Integer
Private mvarstrTipoDescripcion As String

Private mvariIdEmpleado As Integer

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

Public Property Let strPsw(ByVal vData As String)
    mvarstrPsw = vData
End Property

Public Property Get strPsw() As String
    strPsw = mvarstrPsw
End Property

Public Property Let strId(ByVal vData As String)
    mvarstrId = vData
End Property

Public Property Get strId() As String
    strId = mvarstrId
End Property

Public Property Let iId(ByVal vData As Integer)
    mvariId = vData
End Property

Public Property Get iId() As Integer
    iId = mvariId
End Property

Public Property Let iEmpleado(ByVal vData As Integer)
    mvariIdEmpleado = vData
End Property

Public Property Get iEmpleado() As Integer
    iEmpleado = mvariIdEmpleado
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    
    iTipo = mvariTipo
    
End Property

Public Property Let strTipoDescripcion(ByVal vData As String)
    mvarstrTipoDescripcion = vData
End Property

Public Property Get strTipoDescripcion() As String
    strTipoDescripcion = mvarstrTipoDescripcion
End Property

Public Property Let strNombre(ByVal vData As String)
    mvarstrNombre = vData
End Property

Public Property Get strNombre() As String
    strNombre = mvarstrNombre
End Property

Public Property Let strAppPaterno(ByVal vData As String)
    mvarstrAppPaterno = vData
End Property

Public Property Get strAppPaterno() As String
    strAppPaterno = mvarstrAppPaterno
End Property

Public Property Let strAppMaterno(ByVal vData As String)
    mvarstrAppMaterno = vData
End Property

Public Property Get strAppMaterno() As String
    strAppMaterno = mvarstrAppMaterno
End Property

Public Function fnPermiteAcceso(strUsuario As String, strClave As String, ByRef bUsrPsw As Boolean) As Boolean
    
    Dim cDatos As New Collection
    Dim Registro As Collection
    Dim oCampo As Campo
    
    fnPermiteAcceso = False
    
    'Call fnObtenConfiguracionIni
    
    Set cDatos = fnObtenInformacion(strUsuario, gAlmacen)
    
    If cDatos.Count > 0 Then 'Si encontró el Usuario
        
        Set Registro = cDatos(1)
        'Obten el Psw
        Set oCampo = Registro(2)
        
        If strClave = oCampo.Valor Then 'Verifica si es la clave dada
            
            'Obten el id del empleado
            'Set oCampo = Registro(1)
            
            'Set oUsuario = New Usuario
            'oUsuario.oBD = Me
            mvarstrId = strUsuario
            mvarstrPsw = strClave
            
            'mvariIdEmpleado = oCampo.Valor
            
            Set oCampo = Registro(3)
            mvariTipo = oCampo.Valor
            
            'Set oCampo = Registro(5)
            'mvarstrNombre = oCampo.Valor
            'Set oCampo = Registro(6)
            'mvarstrAppPaterno = oCampo.Valor
            'Set oCampo = Registro(7)
            'mvarstrAppMaterno = oCampo.Valor
            
            'Obten los permisos del usuario
            'fnPermisos
            
            'registra actividad
            oBitacora.Nombre = NOMBRE_SOLUCION & "-SISTEMA"
            oBitacora.Usuario = NOMBRE_SOLUCION
            oBitacora.registra ("SE HA REGISTRADO EL USUARIO : " + strUsuario)
            
            fnPermiteAcceso = True
        Else
            bUsrPsw = False
        End If
     
    Else
        bUsrPsw = True
    End If
        
End Function

Private Function fnObtenInformacion(strUsuario As String, iSalon As Integer) As Collection

    Dim oBD As New DataBase
    
    oBD.Procedimiento = "USUARIOSBuscar"
    'oBD.Parametros = "'" & strUsuario & "', " & iSalon
    oBD.Parametros = "'" & strUsuario & "'"
    Set m_cDatos = oBD.fnObtenInformacion()
    Set fnObtenInformacion = m_cDatos
    Set oBD = Nothing
    
End Function
    

Public Function fnPermisos() As Collection
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "obtenAccesos"
    oBD.Parametros = gAlmacen & ", " & mvariIdEmpleado
    
    Set m_cDatos = oBD.fnObtenInformacion()
    
    Set fnPermisos = m_cDatos
    
    Set oBD = Nothing
    
End Function

Public Function catalogoPerfiles() As Boolean
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "TIPO_USUARIOSelProc"
    oBD.Parametros = ""
    
    Set m_cDatos = oBD.fnObtenInformacion()
    
    catalogoPerfiles = oBD.bDatos
    Set oBD = Nothing
End Function

Public Function catalogoUsuarios() As Boolean
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "USUARIOCatalogo"
    oBD.Parametros = ""
    
    Set m_cDatos = oBD.fnObtenInformacion()
    
    catalogoUsuarios = oBD.bDatos
    Set oBD = Nothing
End Function

Public Function actualizaClave(iEmpleado As Integer, _
                               strClaveNueva As String, _
                               iTipoUsuario As Integer, iClave As Integer)

    Dim oBD As New DataBase
    oBD.Procedimiento = "USUARIOUpd"
    oBD.Parametros = iEmpleado & ", '" & strClaveNueva & "', " & iTipoUsuario & ", " & iClave
    oBD.bDatos = False
    oBD.bFolio = False
    oBD.fnInsert
    Set oBD = Nothing
    
End Function

Public Function registra(iTipoUsuario As Integer, _
                          strUsuario As String, _
                          strClave As String)
                          
    Dim oBD As New DataBase
    oBD.Procedimiento = "USUARIOInsProc"
    oBD.Parametros = iTipoUsuario & ", '" & strUsuario & "', '" & strClave & "'"
    oBD.bDatos = False
    oBD.bFolio = False
    oBD.fnInsert
    Set oBD = Nothing
    
End Function

Public Function actualiza(iSalon As Integer, _
                          iEmpleado As Integer, _
                          iTipoUsuario As Integer, _
                          strClave As String)
                          
    Dim oBD As New DataBase
    oBD.Procedimiento = "USUARIOUpdProc"
    oBD.Parametros = iSalon & ", " & iEmpleado & ", " & iTipoUsuario & ", '" & strClave & "'"
    oBD.bDatos = False
    oBD.bFolio = False
    oBD.fnInsert
    Set oBD = Nothing
                          
End Function

Public Function elimina(iSalon As Integer, _
                        iEmpleado As Integer)
                          
    Dim oBD As New DataBase
    oBD.Procedimiento = "USUARIODelProc"
    oBD.Parametros = iSalon & ", " & iEmpleado
    oBD.bDatos = False
    oBD.bFolio = False
    oBD.fnInsert
    Set oBD = Nothing
End Function

Public Function busca(iEmpleado As Integer) As Boolean
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "USUARIOSelProc"
    oBD.Parametros = iEmpleado
    
    Set m_cDatos = oBD.fnObtenInformacion()
    
    busca = oBD.bDatos
    Set oBD = Nothing
    
End Function


Public Function buscaUsuario(strUsuario As String) As Boolean
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "USUARIOSelProcPorUsuario"
    oBD.Parametros = "'" & strUsuario & "'"
    
    Set m_cDatos = oBD.fnObtenInformacion()
    
    m_bDatos = oBD.bDatos
    buscaUsuario = oBD.bDatos
    
    Set oBD = Nothing
    
End Function

Public Function catalogoFunciones() As Boolean
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "FUNCIONSelProc"
    oBD.Parametros = ""
    
    Set m_cDatos = oBD.fnObtenInformacion()
    
    catalogoFunciones = oBD.bDatos
    Set oBD = Nothing
End Function

Public Function registraPerfil(strPerfil As String) As Integer

    Dim iPerfil As Integer
    Dim cRegistros As New Collection
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    Dim oBD As New DataBase
    
    cRegistro.Add oCampo.CreaCampo(adInteger, , , strPerfil)
    cRegistros.Add cRegistro
    
    oBD.Procedimiento = "TIPO_USUARIOInsProc"
    oBD.Parametros = ""
    oBD.bFolio = True
    oBD.cDetalle = cRegistros
    iPerfil = oBD.fnInsert
    Set oBD = Nothing
    registraPerfil = iPerfil
    
'    If iPerfil > 0 Then
'
'        m_oBD.Procedimiento = "PERFILInsProc"
'        m_oBD.Parametros = ""
'        m_oBD.bFolio = False
'        m_oBD.cDetalle = cFunciones
'
'        Call m_oBD.fnInsertaDetalle
'
'    End If
    
End Function

Public Function actualizaPerfil(iPerfil As Integer, cFunciones As Collection) As Boolean

    Dim oBD As New DataBase
    oBD.Procedimiento = "PERFILDelProc"
    oBD.Parametros = iPerfil
    oBD.bDatos = False
    oBD.bFolio = False
    oBD.fnInsert

    oBD.Procedimiento = "PERFILInsProc"
    oBD.Parametros = ""
    oBD.cDetalle = cFunciones
    
    Call oBD.fnInsertaDetalle
    Set oBD = Nothing
    
End Function

Public Function obtenFuncionesPerfil(iPerfil As Integer) As Boolean
    
    Dim oBD As New DataBase
    oBD.Procedimiento = "PERFILSelProc"
    oBD.Parametros = iPerfil
    
    Set m_cDatos = oBD.fnObtenInformacion()
    
    obtenFuncionesPerfil = oBD.bDatos
    Set oBD = Nothing
    
End Function

