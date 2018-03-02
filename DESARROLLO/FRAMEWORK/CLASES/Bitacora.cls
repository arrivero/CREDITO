VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E7390F8033E"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Attribute VB_Ext_KEY = "Member0" ,"Usuario"
Option Explicit

'##ModelId=3E7390FA0142
Private m_iNumeroArchivo As Integer
'##ModelId=3E7390FA0156
Private m_strNombreArchivo As String

'##ModelId=3E7390FA016A
Private mvarstrUsuario As String
'##ModelId=3E7390FA017F
Private mvarstrNombre As String

'##ModelId=3E7390FA0193
Public Property Let Usuario(ByVal vData As String)
    mvarstrUsuario = vData
End Property

'##ModelId=3E7390FA01A8
Public Property Let Nombre(ByVal vData As String)
    mvarstrNombre = vData
End Property

'##ModelId=3E7390FA01C5
Private Sub Class_Initialize()

    Dim DirSistema As String
    Dim dtFecha As Date
    
    'ESTA INICIALIZACION ES POR CONFIGURACION
    DirSistema = App.Path + "\seguimiento\"
    dtFecha = Date
    m_strNombreArchivo = DirSistema + NOMBRE_SOLUCION + "Bitacora" & DatePart("yyyy", dtFecha) & DatePart("m", dtFecha) & DatePart("d", dtFecha) & ".bit"
    m_iNumeroArchivo = FreeFile
    Open m_strNombreArchivo For Output As m_iNumeroArchivo
    
End Sub

'##ModelId=3E7390FA01CF
Private Sub Class_Terminate()

    Close m_iNumeroArchivo

End Sub

'##ModelId=3E7390FA01D9
Public Function registra(ByVal strAccion As String) As Variant

    Dim dtFecha
    
    dtFecha = Date
    
    Print #m_iNumeroArchivo, dtFecha & " Usr: " & mvarstrUsuario & " Nombre: " & mvarstrNombre & " Actividad: " & strAccion
        
End Function


