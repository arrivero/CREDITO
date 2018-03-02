VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3E7390F40176"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private oConeccion As ADODB.Connection

Private mvarstrSP As String
Private mvarstrParametros As String
Private m_bFolio As Boolean

Private m_bDatos As Boolean
Private m_cRegistros As Collection

Public Property Get cDetalle() As Collection
    Set cDetalle = m_cRegistros
End Property

Public Property Let cDetalle(cRegistros As Collection)
    Set m_cRegistros = cRegistros
End Property

Public Property Get bFolio() As Boolean
    bFolio = m_bFolio
End Property

Public Property Let bFolio(ByVal vData As Boolean)
    m_bFolio = vData
End Property

Public Property Get bDatos() As Boolean
    bDatos = m_bDatos
End Property

Public Property Let bDatos(ByVal vData As Boolean)
    m_bDatos = vData
End Property

Public Property Let Procedimiento(ByVal vData As String)
    mvarstrSP = vData
End Property

Public Property Get Procedimiento() As String
    Procedimiento = mvarstrSP
End Property

Public Property Let Parametros(ByVal vData As String)
    mvarstrParametros = vData
End Property


Private Sub Class_Initialize()

'    strServidor = GetSetting(NOMBRE_SOLUCION, "ConfAccesoDB", "SERVIDOR", NOMBRE_EQUIPO)
'    If strServidor = "0" Or strServidor <> NOMBRE_EQUIPO Then
'        SaveSetting NOMBRE_SOLUCION, "ConfAccesoDB", "SERVIDOR", NOMBRE_EQUIPO
'        SaveSetting NOMBRE_SOLUCION, "ConfAccesoDB", "BASEDATOS", NOMBRE_BASEDATOS
'    Else
'        strBaseDatos = GetSetting(NOMBRE_SOLUCION, "ConfAccesoDB", "BASEDATOS", NOMBRE_BASEDATOS)
'    End If

    'Open a connection.
    'Set oConeccion = ObtenConeccion(False)
    
    'If oConeccion.State = adStateOpen Then
    
    '    Dim oComando As New ADODB.Command
    '    Set oComando.ActiveConnection = oConeccion
    
        'DEFINE EL FORMATO DE LA FECHA DEL MANEJADOR DE BASE DE DATOS
    '    oComando.CommandText = "SET DATEFORMAT 'mdy'"
    '    Call oComando.Execute
        
    '    Set oComando = Nothing
        
    'End If
    
    'oConeccion.Close
    'Set oConeccion = Nothing
    
End Sub

Private Function ObtenConeccion(bServerSide As Boolean) As ADODB.Connection

    Dim oConeccion As New ADODB.Connection
    Dim sCnStr As String

    oConeccion.Provider = "sqloledb"
    'oConeccion.CommandTimeout = 0

    'sCnStr = "Server=ABRAHAMRIVERO;Database=AUSSIE;Trusted_Connection=yes"
    sCnStr = "Server=" & strServidor & ";Database=" & strBaseDatos & ";Trusted_Connection=yes"

    oConeccion.Open sCnStr

    Set ObtenConeccion = oConeccion
    
    
'    Dim oConeccion As New ADODB.Connection
'    Dim sCnStr As String
'
'    oConeccion.Provider = "sqloledb"
'    'oConeccion.CommandTimeout = 0
'
'    'sCnStr = "Server=" & strServidor & ";Database=" & strBaseDatos & ";Trusted_Connection=yes"
'    sCnStr = "Server=" & strServidor & ";Database=" & strBaseDatos
'
'    oConeccion.Open sCnStr, "Abraham", "rivero"
'    'oConeccion.Open sCnStr
'
'    Set ObtenConeccion = oConeccion
    
End Function

Private Function fnEjecuta(oComando As ADODB.Command) As ADODB.Recordset

    oComando.CommandText = "EXEC " & mvarstrSP & " " & mvarstrParametros
    oComando.CommandTimeout = 0
    Set fnEjecuta = oComando.Execute
    
End Function

Public Function fnInsert() As Long
    
    On Error GoTo VerErrores
    
    If m_bFolio Then
        fnInsert = fnInsertPadre()
    Else
        Set oConeccion = ObtenConeccion(False)
        
        If oConeccion.State = adStateOpen Then
        
            Dim oComando As New ADODB.Command
            oComando.ActiveConnection = oConeccion
            oComando.CommandTimeout = 0
            
            oComando.CommandText = "EXEC " & mvarstrSP & " " & mvarstrParametros
            
            oComando.Execute
            
            Set oComando = Nothing
            
            oConeccion.Close
            Set oConeccion = Nothing
            
        End If
            
    End If

GoTo cierre
    
VerErrores:

    Dim errLoop As ADODB.Error
    Dim strError As String
    
    For Each errLoop In oConeccion.Errors
    
       strError = "Error Number: " & errLoop.Number
       strError = strError & "  Description: " & errLoop.Description
       strError = strError & "  Source: " & errLoop.Source
       strError = strError & "  SQL State: " & errLoop.SQLState
       strError = strError & "  Native Error: " & errLoop.NativeError
    
        MsgBox strError
                
    Next
    
    oConeccion.Errors.Clear
    
cierre:
        'oConeccion.Close
        'Set oConeccion = Nothing

End Function

Private Function fnInsertPadre() As Long

    On Error GoTo VerErrores
    
    
    Set oConeccion = ObtenConeccion(False)
    
    If oConeccion.State = adStateOpen Then
    
        Dim oComando As New ADODB.Command
        Dim rsConsulta As New ADODB.Recordset
        Dim param As ADODB.Parameter
        
        Set oComando.ActiveConnection = oConeccion
        
        oComando.CommandTimeout = 0
        oComando.CommandText = mvarstrSP
        oComando.CommandType = adCmdStoredProc
        
        oComando.Parameters.Refresh
        Dim iParametros As Integer
        
        iParametros = oComando.Parameters.Count
        
        Dim Registro As Collection
        Dim oCampo As Campo
        Dim i As Integer
        i = 2
        For Each Registro In m_cRegistros
            For Each oCampo In Registro
                oComando.Parameters(i).Value = oCampo.Valor
                i = i + 1
            Next oCampo
        
            Set rsConsulta = oComando.Execute
        
        Next Registro
        
        If Not rsConsulta.EOF And Not rsConsulta.BOF Then
            Debug.Print rsConsulta(0)
            rsConsulta.Close
        End If
        
        Debug.Print oComando(0)
        Debug.Print oComando(1)
        fnInsertPadre = oComando(1)
        
        Set oComando = Nothing
    End If
    
GoTo cierre

VerErrores:

    Dim errLoop As ADODB.Error
    Dim strError As String
    
    For Each errLoop In oConeccion.Errors
    
       strError = "Error Number: " & errLoop.Number
       strError = strError & "  Description: " & errLoop.Description
       strError = strError & "  Source: " & errLoop.Source
       strError = strError & "  SQL State: " & errLoop.SQLState
       strError = strError & "  Native Error: " & errLoop.NativeError
    
        MsgBox strError
                
    Next
    
    oConeccion.Errors.Clear

cierre:

    oConeccion.Close
    Set oConeccion = Nothing

End Function

Public Function fnInsertaDetalle() As Integer

    On Error GoTo VerErrores
    
    fnInsertaDetalle = 0
    
    Set oConeccion = ObtenConeccion(False)
    
    If oConeccion.State = adStateOpen Then
    
        Dim oComando As New ADODB.Command
        
        Set oComando.ActiveConnection = oConeccion
        
        oComando.CommandTimeout = 0
        oComando.CommandText = mvarstrSP
        oComando.CommandType = adCmdStoredProc
        
        oComando.Parameters.Refresh
        
        Dim Registro As Collection
        Dim i As Integer
        
        Dim oCampo As Campo
        For Each Registro In m_cRegistros
            i = 1
            For Each oCampo In Registro
               oComando.Parameters(i).Value = oCampo.Valor
               i = i + 1
            Next oCampo
            
            oComando.Execute
            
        Next Registro
        
        Set oComando = Nothing
        
    End If
    
GoTo cierre

VerErrores:

   fnInsertaDetalle = 1
   
    Dim errLoop As ADODB.Error
    Dim strError As String
    
    For Each errLoop In oConeccion.Errors
    
       strError = "Error Number: " & errLoop.Number
       strError = strError & "  Description: " & errLoop.Description
       strError = strError & "  Source: " & errLoop.Source
       strError = strError & "  SQL State: " & errLoop.SQLState
       strError = strError & "  Native Error: " & errLoop.NativeError
    
        MsgBox strError
                
    Next
    
    oConeccion.Errors.Clear

cierre:

    oConeccion.Close
    Set oConeccion = Nothing

End Function

Public Function fnObtenInformacionEnArchivo(strArchivo As String) As String
    
    Dim cDatos As New Collection
    
    Set oConeccion = ObtenConeccion(False)
    If oConeccion.State = adStateOpen Then
    
        Dim oComando As New ADODB.Command
        Dim rsConsulta As ADODB.Recordset
        
        Set oComando.ActiveConnection = oConeccion
        Set rsConsulta = fnEjecuta(oComando)
        
        m_bDatos = False
        
        If rsConsulta.EOF <> True Then
                    
            Dim iNumeroArchivo As Integer

            'Abrir archivo para escritura
            If abreArchivofn(strArchivo, iNumeroArchivo, PARA_ESCRITURA) = True Then
            
                Dim iColumna As Integer
                Dim strRegistro As String

                While rsConsulta.EOF <> True
                    
                    For iColumna = 0 To rsConsulta.Fields.Count - 1
                    
                        If iColumna = 0 Then
                            strRegistro = rsConsulta.Fields(iColumna)
                        Else
                        
                            If iColumna = 2 Or iColumna = 5 Then '2 - Fecha (segun es la fecha de pago, para que sirve esta?) o 5 - fecha de inicio de cobro del credito
                                
                                strRegistro = strRegistro & "," & Chr(34) & Format(rsConsulta.Fields(iColumna), "yyyy/mm/dd") & Chr(34)
                                
                            Else 'es 1 (factura) o 3 (suma de pagos, o sea lo pagado)
                            
                                strRegistro = strRegistro & "," & CStr(rsConsulta.Fields(iColumna))
                                
                            End If
                            
                        End If
                    
                        m_bDatos = True
                        
                    Next iColumna
                        
                    agregaRegistroArchivofn strRegistro, iNumeroArchivo
                    
                    rsConsulta.MoveNext
                        
                Wend
                                
                'Enviar mensaje de proceso terminado
                fnObtenInformacionEnArchivo = "Los estados de cuenta estan listos"
                
                'Cerrar el archivo
                cierraArchivofn iNumeroArchivo
            Else
            
                'Enviar mensaje de proceso terminado
                fnObtenInformacionEnArchivo = "No hubo estados de cuenta."
            
            End If
            
        Else
        
            'Enviar mensaje de proceso terminado
            fnObtenInformacionEnArchivo = "Problemas para almacenar los estados de cuentas, verifique con su administrador."
        
        End If
        
        rsConsulta.Close
        Set rsConsulta = Nothing
        
        Set oComando = Nothing
        
    End If
    
    oConeccion.Close
    Set oConeccion = Nothing
    
End Function

Public Function fnObtenInformacionEnArchivoUno(strArchivo As String) As String
    
    Dim cDatos As New Collection
    
    Set oConeccion = ObtenConeccion(False)
    If oConeccion.State = adStateOpen Then
    
        Dim oComando As New ADODB.Command
        Dim rsConsulta As ADODB.Recordset
        
        Set oComando.ActiveConnection = oConeccion
        Set rsConsulta = fnEjecuta(oComando)
        
        m_bDatos = False
        
        If rsConsulta.EOF <> True Then
                    
            Dim iNumeroArchivo As Integer

            'Abrir archivo para escritura
            If abreArchivofn(strArchivo, iNumeroArchivo, PARA_ESCRITURA) = True Then
            
                Dim iColumna As Integer
                Dim strRegistro As String

                While rsConsulta.EOF <> True
                    
                    For iColumna = 0 To rsConsulta.Fields.Count - 1
                    
                        If iColumna = 0 Then
                            strRegistro = rsConsulta.Fields(iColumna)
                        Else
                        
                            strRegistro = strRegistro & "," & CStr(rsConsulta.Fields(iColumna))
                            
                        End If
                    
                        m_bDatos = True
                        
                    Next iColumna
                        
                    agregaRegistroArchivofn strRegistro, iNumeroArchivo
                    
                    rsConsulta.MoveNext
                        
                Wend
                                
                'Enviar mensaje de proceso terminado
                fnObtenInformacionEnArchivoUno = "Los estados de cuenta estan listos"
                
                'Cerrar el archivo
                cierraArchivofn iNumeroArchivo
            Else
            
                'Enviar mensaje de proceso terminado
                fnObtenInformacionEnArchivoUno = "No hubo estados de cuenta."
            
            End If
            
        Else
        
            'Enviar mensaje de proceso terminado
            fnObtenInformacionEnArchivoUno = "Problemas para almacenar los estados de cuentas, verifique con su administrador."
        
        End If
        
        rsConsulta.Close
        Set rsConsulta = Nothing
        
        Set oComando = Nothing
        
    End If
    
    oConeccion.Close
    Set oConeccion = Nothing
    
End Function

Public Function fnObtenInformacion() As Collection
    
    Dim cDatos As New Collection
    
    Set oConeccion = ObtenConeccion(False)
    If oConeccion.State = adStateOpen Then
    
        Dim oComando As New ADODB.Command
        Dim rsConsulta As ADODB.Recordset
        
        Set oComando.ActiveConnection = oConeccion
        'oComando.CommandTimeout = 0
        
        Set rsConsulta = fnEjecuta(oComando)
        
        Dim oCampo As New Campo
        Dim cRegistro As Collection
        
        m_bDatos = False
        Dim iColumna As Integer
            While rsConsulta.EOF <> True
                
                Set cRegistro = New Collection
                
                For iColumna = 0 To rsConsulta.Fields.Count - 1
                    cRegistro.Add oCampo.CreaCampo(adInteger, , , rsConsulta.Fields(iColumna))
                    m_bDatos = True
                Next iColumna
                
                cDatos.Add cRegistro
                
                rsConsulta.MoveNext
                
            Wend
        
        Set m_cRegistros = cDatos
        
        rsConsulta.Close
        Set rsConsulta = Nothing
        
        Set oComando = Nothing
        
    End If
    
    oConeccion.Close
    Set oConeccion = Nothing
    
    Set fnObtenInformacion = cDatos
    
End Function

Public Function fnGeneraInformacion(ByRef arrDatos() As Variant) As Boolean

    Set oConeccion = ObtenConeccion(False)
    If oConeccion.State = adStateOpen Then

        Dim oComando As New ADODB.Command
        Dim rsConsulta As ADODB.Recordset

        Set oComando.ActiveConnection = oConeccion
        Set rsConsulta = fnEjecuta(oComando)

        Dim arrGen() As Variant
        Dim iRenglon As Integer
        Dim iArreglos As Integer
        Dim iColumna As Integer
        Dim iRenglones As Integer
        fnGeneraInformacion = False

        iRenglon = 1
        While rsConsulta.EOF <> True
            iRenglones = rsConsulta.RecordCount
            For iColumna = 0 To rsConsulta.Fields.Count - 1
                arrGen = arrDatos(iColumna + 1)
                ReDim Preserve arrGen(iRenglon)
                arrGen(iRenglon) = rsConsulta.Fields(iColumna)
                arrDatos(iColumna + 1) = arrGen
                fnGeneraInformacion = True
            Next iColumna

            iRenglon = iRenglon + 1

            rsConsulta.MoveNext

        Wend

        rsConsulta.Close
        Set rsConsulta = Nothing

        Set oComando = Nothing

    End If

    oConeccion.Close
    Set oConeccion = Nothing

End Function
