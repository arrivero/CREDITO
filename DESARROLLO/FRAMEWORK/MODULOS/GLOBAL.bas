Attribute VB_Name = "Global"
Option Explicit

Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal strVerb As String, _
                                                     ByVal strAplicacion As String, ByVal strParametros As String, _
                                                     ByVal strDirectorio As String, ByVal iPresenta As Integer) As Long

'Objeto que contiene la coneccion y acceso a la base de datos
'Public oAlmacen As New Almacen
'Public oBitacora As New Bitacora
'Public oUsuario As New Usuario

Global Const Valor = 0
Global Const CADENA = 1
Global Const fecha = 2
Global Const BOOL = 3
Global Const DEC = 4

Global Const alta = 0
Global Const BAJA = 1
Global Const CAMBIO = 2

Global Const Aceptar = 1
Global Const CANCELAR = 0

'CONSTANTES PARA MANEJO DE ARCHIVOS
Global Const PARA_ESCRITURA = 0
Global Const PARA_LECTURA = 1

'CONSTANTES PARA LA CLASE cMes y el CONTROL SSMonth
Global Const DOMINGO = 1
Global Const LUNES = 2
Global Const MARTES = 3
Global Const MIERCOLES = 4
Global Const JUEVES = 5
Global Const VIERNES = 6
Global Const SABADO = 7

'CONSTANTES DE TIPO DE USUARIO
Global Const USUARIO_GERENTE = 1
Global Const USUARIO_ADMINSTRADOR = 2
Global Const USUARIO_USUARIO = 3

'CONSTANTES DE CONFIGURACION
Global Const CONF_HAY_MODULO_CONTABLE = 1
Global Const CONF_IVA = 2


Public gNombre As Variant
Public DirSys As Variant
Public dirsys1 As String
Public gAlmacen As Integer
Public gPrintPed As Variant

Public strServidor As Variant
Public strBaseDatos As Variant

Public gstrHoraInicial As String
Public gstrHoraFinal As String
Public gstrHoraInicialFinal As String
Public gstrReporteEnPantalla As String
Public gstrMandaImprimir As String
Public gEncriptado As String

Global cPagosMem As New Collection

Global oBitacora As New Bitacora

Global gstrUsuario As String
Global giTipoUsuario As Integer

Private arrUnidad(0 To 19) As String

Private arrDecena(0 To 9) As String

Private arrCentena(1 To 9) As String

Private Sub InicializarArrays()
    
    'Asignar los valores
    arrUnidad(1) = ""
    arrUnidad(1) = "UN"
    arrUnidad(2) = "DOS"
    arrUnidad(3) = "TRES"
    arrUnidad(4) = "CUATRO"
    arrUnidad(5) = "CINCO"
    arrUnidad(6) = "SEIS"
    arrUnidad(7) = "SIETE"
    arrUnidad(8) = "OCHO"
    arrUnidad(9) = "NUEVE"
    arrUnidad(10) = "DIEZ"
    arrUnidad(11) = "ONCE"
    arrUnidad(12) = "DOCE"
    arrUnidad(13) = "TRECE"
    arrUnidad(14) = "CATORCE"
    arrUnidad(15) = "QUINCE"
    arrUnidad(16) = "DIEZ Y SEIS"
    arrUnidad(17) = "DIEZ Y SIETE"
    arrUnidad(18) = "DIEZ Y OCHO"
    arrUnidad(19) = "DIEZ Y NUEVE"
    
    '
    arrDecena(0) = ""
    arrDecena(1) = "DIEZ"
    arrDecena(2) = "VEINTE"
    arrDecena(3) = "TREINTA"
    arrDecena(4) = "CUARENTA"
    arrDecena(5) = "CINCUENTA"
    arrDecena(6) = "SESENTA"
    arrDecena(7) = "SETENTA"
    arrDecena(8) = "OCHENTA"
    arrDecena(9) = "NOVENTA"
    '
    arrCentena(1) = "CIEN"
    arrCentena(2) = "DOSCIENTOS"
    arrCentena(3) = "TRESCIENTOS"
    arrCentena(4) = "CUATROCIENTOS"
    arrCentena(5) = "QUINIENTOS"
    arrCentena(6) = "SEISCIENTOS"
    arrCentena(7) = "SETECIENTOS"
    arrCentena(8) = "OCHOCIENTOS"
    arrCentena(9) = "NOVECIENTOS"

End Sub

Public Function convierteMontoConLentra(strNo As String) As String
    
    Dim strNoEntero, strDecimal As String
    Dim strSubEntero As String
    Dim strMillon As String
    Dim strCantidadConLetra, strMontoConLetra, strCantidadConLetraTemp As String
    Dim iPos, iBloque As Integer
    Dim bMiles, bUnidades As Boolean
    
    InicializarArrays
    
    iPos = InStr(1, strNo, ".")
    If iPos = 0 Then
        strNoEntero = strNo
        strDecimal = Format("0", "00")
    Else
        strNoEntero = Mid(strNo, 1, iPos - 1)
        strDecimal = Mid(strNo, iPos + 1, 2)
    End If
    
    iBloque = 1
    
    Do
        strSubEntero = Right(strNoEntero, 3)
            
        strCantidadConLetraTemp = obtenCantidadLetra(strSubEntero)
                
        Select Case iBloque
            Case Is = 1 'cienes
                strCantidadConLetra = strCantidadConLetraTemp
                If Trim(strCantidadConLetraTemp) <> "" Then
                    bUnidades = True
                Else
                    bUnidades = False
                End If
            Case Is = 2 'miles
                If Trim(strCantidadConLetraTemp) <> "" Then
                    strCantidadConLetra = strCantidadConLetraTemp & " MIL " & strCantidadConLetra
                    bMiles = True
                Else
                    bMiles = False
                End If
            Case Is = 3 'millones
                If Trim(strCantidadConLetraTemp) <> "" Then
                    If Trim(strCantidadConLetraTemp) = "UN" Then
                        If bUnidades = True Then
                            strMillon = " MILLON "
                        Else
                            If bMiles = True Then
                                strMillon = " MILLON "
                            Else
                                strMillon = " MILLON DE "
                            End If
                        End If
                    Else
                        If bUnidades = True Then
                            strMillon = " MILLONES "
                        Else
                            If bMiles = True Then
                                strMillon = " MILLONES "
                            Else
                                strMillon = " MILLONES DE "
                            End If
                        End If
                    End If
                    strCantidadConLetra = strCantidadConLetraTemp & strMillon & strCantidadConLetra
                End If
            Case Is = 4 'miles de millones
                strCantidadConLetra = strCantidadConLetraTemp & " MIL " & strCantidadConLetra
            Case Is = 5 'billones
                strCantidadConLetra = strCantidadConLetraTemp & " BILLONES " & strCantidadConLetra
        End Select
        
        iBloque = iBloque + 1
        
        If Len(strNoEntero) - 3 > 0 Then
            strNoEntero = Mid(strNoEntero, 1, Len(strNoEntero) - 3)
        Else
            Exit Do
        End If
        
    Loop While Len(strNoEntero) > 0
    
    If strCantidadConLetra <> "" Then
        Dim strPesos As String
        If strCantidadConLetra = "UN" Then
            strPesos = " PESO "
        Else
            strPesos = " PESOS "
        End If
        
        convierteMontoConLentra = "(" & strCantidadConLetra & strPesos & strDecimal & "/100 M.N.)"
    Else
        convierteMontoConLentra = "(CERO PESOS" & strDecimal & "/100 M.N.)"
    End If

End Function

Private Function obtenCantidadLetra(strSubEntero As String) As String
    
    Dim strSubSubEntero, strCentena, strDecena, strUnidad As String
    Dim ivalorCentena, ivalorDecena, ivalorUnidad As Integer
    
    If Val(strSubEntero) > 99 Then
        ivalorCentena = Val(Left(strSubEntero, 1))
        If ivalorCentena = 1 Then
            ivalorDecena = Val(Right(strSubEntero, 2))
            If ivalorDecena > 0 Then
                strCentena = "CIENTO"
            Else
                ivalorCentena = Val(Left(strSubEntero, 1))
                strCentena = arrCentena(ivalorCentena)
            End If
        Else
            ivalorCentena = Val(Left(strSubEntero, 1))
            strCentena = arrCentena(ivalorCentena)
        End If
        
        strSubSubEntero = Right(strSubEntero, 2)
        
        If Val(strSubSubEntero) < 20 Then
            strUnidad = arrUnidad(Val(strSubSubEntero))
        Else
            strDecena = arrDecena(Val(Left(strSubSubEntero, 1)))
            If Val(Right(strSubSubEntero, 1)) > 0 Then
                strUnidad = "Y "
                strUnidad = strUnidad & arrUnidad(Val(Right(strSubSubEntero, 1)))
            Else
                strUnidad = arrUnidad(Val(Right(strSubSubEntero, 1)))
            End If
        End If
        
    Else
    
        If Val(strSubEntero) < 20 Then
            strUnidad = arrUnidad(Val(strSubEntero))
        Else
            ivalorDecena = Val(Left(strSubEntero, 1))
            ivalorUnidad = Val(Right(strSubEntero, 1))
            strDecena = arrDecena(ivalorDecena)
            If ivalorUnidad > 0 Then
                strUnidad = "Y "
                strUnidad = strUnidad & arrUnidad(ivalorUnidad)
            Else
                strUnidad = arrUnidad(ivalorUnidad)
            End If
        End If
        
    End If
    
    If strCentena <> "" Then
        If strDecena <> "" Then
            If strUnidad <> "" Then
                obtenCantidadLetra = strCentena & " " & strDecena & " " & strUnidad
            Else
                obtenCantidadLetra = strCentena & " " & strDecena
            End If
        Else
            If strUnidad <> "" Then
                obtenCantidadLetra = strCentena & " " & strUnidad
            Else
                obtenCantidadLetra = strCentena
            End If
        End If
    Else
    
        If strDecena <> "" Then
            If strUnidad <> "" Then
                obtenCantidadLetra = strDecena & " " & strUnidad
            Else
                obtenCantidadLetra = strDecena
            End If
        Else
            If strUnidad <> "" Then
                obtenCantidadLetra = strUnidad
            Else
                obtenCantidadLetra = ""
            End If
        End If
    
    End If
    
End Function

Public Sub fnObtenConfiguracionIni()
    
    Dim Todo As String
    Dim XI As String
    Dim xlabel As String
    Dim xvalor As String
    
    Open App.Path & "\" & NOMBRE_FILE_CONFIG For Input As 1
    
    Do While Not EOF(1)
        
        Line Input #1, Todo
        
        XI = InStr(1, Todo, "=")
        xlabel = LTrim(RTrim(Left(Todo, XI - 1)))
        xvalor = LTrim(RTrim(Mid(Todo, XI + 1, 50)))
        
        If xlabel = "Nombre" Then
            gNombre = xvalor
        End If
        
        If xlabel = "Directorio" Then
            'DirSys = xvalor
            DirSys = App.Path & "\REPORTES\"
        End If
        
        If xlabel = "Almacen" Then
            gAlmacen = xvalor
        End If
        
        If xlabel = "Facturas" Then
            gPrintPed = xvalor
        End If
        
        If xlabel = "ServidorDatos" Then
            strServidor = xvalor
        End If
        
        If xlabel = "BaseDeDatos" Then
            strBaseDatos = xvalor
        End If
        
        If xlabel = "HoraInicioOperaciones" Then
            gstrHoraInicial = xvalor
        End If
        
        If xlabel = "HoraFinOperaciones" Then
            gstrHoraFinal = xvalor
        End If
        
        If xlabel = "HoraInicialFinOperaciones" Then
            gstrHoraInicialFinal = xvalor
        End If
        
        If xlabel = "MandaImprimir" Then
            gstrMandaImprimir = xvalor
        End If
        
        If xlabel = "ReporteEnPantalla" Then
            gstrReporteEnPantalla = xvalor
        End If
        
        If xlabel = "ConEncripcion" Then
            gEncriptado = xvalor
        End If
    Loop
    
    Close #1
    
End Sub

Public Function fnLlenaComboCollecion(cbCombo As ComboBox, cDatos As Collection, iOpcionNuevo, strTexto As String)
    
    Dim oCampo As New Campo
    Dim Registro As Collection
    cbCombo.Clear
    If iOpcionNuevo = 1 Then
        cbCombo.AddItem strTexto
        cbCombo.ItemData(cbCombo.NewIndex) = 0
    End If
    For Each Registro In cDatos
        Set oCampo = Registro(2)
        cbCombo.AddItem (oCampo.Valor)
        Set oCampo = Registro(1)
        cbCombo.ItemData(cbCombo.NewIndex) = oCampo.Valor
    Next Registro

End Function

Public Function fnLimpiaGrid(spGrid As vaSpread) As Variant

    spGrid.Row = spGrid.DataRowCnt
    spGrid.Col = spGrid.DataColCnt
    spGrid.Row2 = -1
    spGrid.Col2 = -1
    spGrid.BlockMode = True
    spGrid.Action = ActionClearText
    spGrid.BlockMode = False
    
End Function

Public Function fnLlenaTablaCollection(Tabla As vaSpread, cDatos As Collection)

    Dim oCampo As New Campo
    Dim Registro As Collection
    Dim lRenglon As Long
    Dim lColumna As Long
    lRenglon = 1
    
    Tabla.MaxRows = cDatos.Count
    
    For Each Registro In cDatos
    
        lColumna = 1
        
        Dim iNumCampos As Integer
        Dim iNumero As Integer
        iNumCampos = Registro.Count
        For Each oCampo In Registro
        
            Tabla.Row = lRenglon
        
            If lColumna <= Tabla.MaxCols Then
                Tabla.Col = lColumna
                Select Case Tabla.CellType
                    Case Is = SS_CELL_TYPE_INTEGER, SS_CELL_TYPE_FLOAT
                        If Not IsNull(oCampo.Valor) Then
                            Tabla.Text = Val(oCampo.Valor)
                        End If
                    Case Is = SS_CELL_TYPE_EDIT, SS_CELL_TYPE_CHECKBOX
                        If Not IsNull(oCampo.Valor) Then
                            Tabla.Text = oCampo.Valor
                        End If
                
                    Case Is = SS_CELL_TYPE_DATE
                        If Not IsNull(oCampo.Valor) Then
                            Dim strFecha As String
                            strFecha = Format(oCampo.Valor, "dd/mm/yyyy")
                            Tabla.Text = strFecha
                        End If
                
                    Case Is = SS_CELL_TIPE_COMBOBOX
                        
                        Dim i As Integer
                        Dim strCadena As String
                        
                        For i = 0 To Tabla.TypeComboBoxCount - 1
                        
                            Tabla.TypeComboBoxIndex = i
                            strCadena = Tabla.TypeComboBoxString
                            If oCampo.Valor = Trim(strCadena) Then
                                Tabla.TypeComboBoxCurSel = i
                                Tabla.TypeComboBoxWidth = 1
                                Exit For
                            End If
                            
                        Next i
                        
                End Select
                
                lColumna = lColumna + 1
                
            End If
            
        Next oCampo
        
        lRenglon = lRenglon + 1
        
    Next Registro

End Function

Public Function periodoObtenStr(ByRef strHoraInicio As String, ByRef strHoraFin As String)

    Dim iPos As Integer
    
    iPos = InStr(1, strHoraInicio, ":", vbTextCompare)
    strHoraInicio = Val(Left(strHoraInicio, iPos + 2))
    
    iPos = InStr(1, strHoraFin, ":", vbTextCompare)
    strHoraFin = Val(Left(strHoraFin, iPos + 2))
    
End Function

Public Function periodoObten(ByVal strHoraInicio As String, ByRef intHoraInicio As Integer, ByVal strHoraFinal As String, ByRef intHoraFin As Integer)

    Dim iPos As Integer
    
    iPos = InStr(1, strHoraInicio, ":", vbTextCompare)
    intHoraInicio = Val(Left(strHoraInicio, iPos - 1))
    
    iPos = InStr(1, strHoraFinal, ":", vbTextCompare)
    intHoraFin = Val(Left(strHoraFinal, iPos - 1))
    
End Function

Public Function fnstrValor(strValor As String) As String

    Dim strPrecio As String
    Dim strPrecioAntesComa As String
    Dim strPrecioDespuesComa As String
    Dim iComaPocision As Integer
    Dim iPosPesos As Integer
    Dim strSigno As String
    
    If InStr(strValor, "-") > 0 Then
        strSigno = "-"
    End If
    
    iPosPesos = InStr(strValor, "$")
    strPrecio = Mid(strValor, iPosPesos + 1, Len(strValor)) 'Quita el $
    
    iPosPesos = InStr(strPrecio, "_")
    strPrecio = Mid(strPrecio, iPosPesos + 1, Len(strPrecio)) 'Quita el - bajo
    
    iComaPocision = InStr(1, strPrecio, ",")
    While iComaPocision > 0
        strPrecioAntesComa = Mid(strPrecio, 1, iComaPocision - 1)
        strPrecioDespuesComa = Mid(strPrecio, iComaPocision + 1, Len(strPrecio))
        strPrecio = strPrecioAntesComa + strPrecioDespuesComa
        iComaPocision = InStr(1, strPrecio, ",")
    Wend
    
    fnstrValor = strSigno + strPrecio

End Function

Public Function fnValidaPeriodoFechas(ByVal FechaInicial As String, ByVal FechaFinal As String, ByRef iFechaErronea As Integer) As String

    Dim strMensaje As String
    
    If Not IsDate(FechaInicial) Then
        strMensaje = "Fecha incial no válida."
        iFechaErronea = 1
        fnValidaPeriodoFechas = strMensaje
        Exit Function
    End If
    If Not IsDate(FechaFinal) Then
        strMensaje = "Fecha final no válida."
        iFechaErronea = 2
        fnValidaPeriodoFechas = strMensaje
        Exit Function
    End If
    
    If DateDiff("d", FechaFinal, FechaInicial) >= 0 Then
        strMensaje = "La fecha final no puede ser menor o igual a la fecha inicial."
        iFechaErronea = 2
        fnValidaPeriodoFechas = strMensaje
        Exit Function
    End If
    
End Function

Public Function fnBuscaTextoCombo(mvarobCombo As ComboBox, ByVal strTexto As String) As Integer
    
    Dim i As Integer
    Dim strCadena As String
    For i = 0 To mvarobCombo.ListCount - 1
        
        If InStr(1, UCase(strTexto), UCase(mvarobCombo.List(i))) > 0 Then
            mvarobCombo.ListIndex = i
            Exit For
        End If
        
    Next i
    
End Function

Public Function fnBuscaElemento(mvarobCombo As ComboBox, ByVal iItemData As Integer) As Integer
    
    Dim i As Integer
    
    For i = 0 To mvarobCombo.ListCount - 1
    
        If iItemData = mvarobCombo.ItemData(i) Then
            mvarobCombo.ListIndex = i
            Exit For
        End If
        
    Next i
    
End Function

Public Function fnBuscaIndiceCombo(cbCombo As ComboBox, iItemData As Integer) As Integer
    Dim i As Integer
    Dim bEncontro As Boolean
    For i = 0 To cbCombo.ListCount - 1
        If iItemData = cbCombo.ItemData(i) Then
            fnBuscaIndiceCombo = i
            Exit For
        End If
    Next i
    
End Function

Public Function fnBuscaIndiceComboEsp(cbCombo As ComboBox, iItemData As Integer, iLlaveEnTexto As Integer)

    Dim i As Integer
    Dim bEncontro As Boolean
    Dim strNumero As String
    Dim iLong As Integer
    
    For i = 0 To cbCombo.ListCount - 1
        If iItemData = cbCombo.ItemData(i) Then
            strNumero = cbCombo.Text
            iLong = InStr(1, strNumero, " ", vbTextCompare)
            strNumero = Mid(strNumero, 1, iLong)
            If Val(strNumero) = iLlaveEnTexto Then
                fnBuscaIndiceComboEsp = i
                Exit For
            End If
        End If
    Next i
    
End Function

Public Function validaFecha(dFecha As Date, bSpread As Boolean) As Boolean
    
    Dim bvalidaFecha As Boolean
    
    If bSpread = True Then
        If 0 = DateDiff("yyyy", Now, CDate(dFecha)) Then
            
            If 0 = DateDiff("m", Now, CDate(dFecha)) Then
                
                If 0 = DateDiff("d", Now, CDate(dFecha)) Then
                    
                    bvalidaFecha = True
                    
                End If
            End If
        End If
    Else
    
        If 0 <= DateDiff("yyyy", Now, CDate(dFecha)) Then
            
            If 0 <= DateDiff("m", Now, CDate(dFecha)) Then
                
                If 0 <= DateDiff("d", Now, CDate(dFecha)) Then
                    
                    bvalidaFecha = True
                    
                End If
            End If
        End If
    End If
    
    validaFecha = bvalidaFecha
    
End Function

Public Function llenaComboSpread(sprGrid As vaSpread, lCol As Long, cDatos As Collection, iConId As Integer)

    Dim strLista As String
    Dim strId As String

    Dim oCampo As New Campo
    Dim Registro As Collection

    sprGrid.Col = lCol
    sprGrid.Row = -1
    sprGrid.Col2 = lCol
    sprGrid.Row2 = -1
    sprGrid.CellType = CellTypeComboBox

    For Each Registro In cDatos

        Set oCampo = Registro(1)
        strId = Format(CStr(oCampo.Valor), "00")

        Set oCampo = Registro(2)
        If iConId = 1 Then
            strLista = strLista + strId + "-" + oCampo.Valor + Chr$(9)
        Else
            strLista = strLista + oCampo.Valor + Chr$(9)
        End If

    Next Registro

    'Quitar el último caracter \t
    If Len(strLista) > 1 Then
        strLista = Mid(strLista, 1, Len(strLista) - 1)
    End If

    sprGrid.TypeComboBoxList = strLista

End Function

Public Function obtenTotalGrid(sprGrid As vaSpread, lCualColumna As Long) As Double
    
    Dim lRow As Long
    Dim dTotal As Double
    
    For lRow = 1 To sprGrid.DataRowCnt
        
        sprGrid.Row = lRow
        sprGrid.Col = lCualColumna
        dTotal = dTotal + Val(fnstrValor(sprGrid.Text))
        
    Next lRow
    
    obtenTotalGrid = dTotal
    
End Function

Public Function existeValorEnGrid(sprGrid As vaSpread, lCualColumna As Long, lValor As Long, lRowCaptura As Long) As Boolean
    
    Dim lRow As Long
    Dim bExiste As Boolean
    
    bExiste = False
    
    For lRow = 1 To sprGrid.DataRowCnt
        
        sprGrid.Row = lRow
        sprGrid.Col = lCualColumna
        If lRowCaptura <> lRow Then
            If lValor = CLng(sprGrid.Text) Then
                bExiste = True
                Exit For
            End If
        End If
    Next lRow
    
    existeValorEnGrid = bExiste
    
End Function

'
'FUNCIONES PARA EL USO Y MANEJO DE ARCHIVOS
'
'

Public Function abreArchivofn(ByVal strNombreArchivo As String, ByRef iNumeroArchivo As Integer, iModo As Integer) As Boolean

    abreArchivofn = True
    
    iNumeroArchivo = FreeFile
    
    Select Case iModo
        Case Is = 0
            Open strNombreArchivo For Output As iNumeroArchivo
            
        Case Is = 1
            Open strNombreArchivo For Input As iNumeroArchivo
            
    End Select
    
End Function

Public Function agregaRegistroArchivofn(strRegistro As String, iNumeroArchivo As Integer)

    'strRegistro = strRegistro + Chr(10) + Chr(13)
    'strRegistro = strRegistro + Chr(13)
    Print #iNumeroArchivo, Mid(strRegistro, 1, Len(strRegistro))

End Function

Public Function obtenRegistrofn(ByVal iNumeroArchivo, ByRef strRegistro)
    Line Input #iNumeroArchivo, strRegistro
End Function

Public Function cierraArchivofn(ByVal iNumeroArchivo)
    Close iNumeroArchivo
End Function

Public Function agregaRegistrosArchivo(iPreciosArchivo As Integer, cDatos As Collection)
                    
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    
    Dim strRegistro As String
    Dim iCampo As Integer
    
    'Por cada producto un registro en el archivo de precios
    For Each cRegistro In cDatos
        
        iCampo = 1
        
        For Each oCampo In cRegistro
        
            If iCampo = 1 Then
                strRegistro = oCampo.Valor
            Else
                strRegistro = strRegistro & ", " & CStr(oCampo.Valor)
            End If
            
            iCampo = iCampo + 1
            
        Next oCampo
        
        agregaRegistroArchivofn strRegistro, iPreciosArchivo
        
    Next cRegistro
    
End Function

Public Function agregaRegistrosArchivoInternet(iPreciosArchivo As Integer, cDatos As Collection)
                    
    Dim cRegistro As Collection
    Dim oCampo As New Campo
    
    Dim strRegistro As String
    Dim iCampo As Integer
    
    'Por cada producto un registro en el archivo de precios
    For Each cRegistro In cDatos
        
        iCampo = 1
        
        For Each oCampo In cRegistro
        
            If iCampo = 1 Then
                strRegistro = oCampo.Valor
            Else
            
                If iCampo = 3 Or iCampo = 6 Then '3 - Fecha (segun es la fecha de pago, para que sirve esta?) o 6 - fecha de inicio de cobro del credito
                
                    
                    strRegistro = strRegistro & "," & Chr(34) & Format(oCampo.Valor, "yyyy/mm/dd") & Chr(34)
                    
                Else 'es 2 (factura) o 4 (suma de pagos, o sea lo pagado)
                
                    strRegistro = strRegistro & "," & CStr(oCampo.Valor)
                    
                End If
                
            End If
            
            iCampo = iCampo + 1
            
        Next oCampo
        
        agregaRegistroArchivofn strRegistro, iPreciosArchivo
        
    Next cRegistro
    
End Function

Public Function buscaEnLista(sprSpread As vaSpread, Col As Long, Valor As Variant)

    Dim lRow As Long
    sprSpread.Col = Col
    For lRow = 1 To sprSpread.DataRowCnt
        sprSpread.Row = lRow
        
        If Valor = Val(sprSpread.Text) Then
            sprSpread.Action = ActionActiveCell
            Exit For
        End If
        
    Next lRow

End Function

Public Function dibujaGrafica(m_cDatos As Collection, grGrafico As VtChart, iPeriodo As Integer) As Boolean

    If m_cDatos.Count > 0 Then

        Dim oCampo As New Campo
        Dim Registro As Collection

        Dim cEtiquetaColumna As New Collection
        Dim cEtiquetasRenglon As New Collection
        Dim seriePeriodo As String

        Dim lColumna As Long
        Dim lRenglon As Long
        Dim lRenglonTemp As Long
        Dim lRenglonGrid As Long
        Dim lColumnaGrid As Long
        
        'lRenglonGrid = 1
        'lColumnaGrid = 1
        'grGrafico.RowCount = 0
        'grGrafico.ColumnCount = 0
        'grGrafico.ColumnLabelCount = 1
        'grGrafico.RowLabelCount = 1

        seriePeriodo = ""

        For Each Registro In m_cDatos

            Set oCampo = Registro(1)

            If seriePeriodo <> oCampo.Valor Then
                
                lRenglon = yaExisteEtiquetaRenglon(oCampo.Valor, cEtiquetasRenglon)
                
                If lRenglon = 0 Then
                    
                    cEtiquetasRenglon.Add oCampo.Valor
                    If cEtiquetasRenglon.Count > grGrafico.RowCount Then
                        grGrafico.RowCount = grGrafico.RowCount + 1
                    End If
                    
                    lRenglonGrid = lRenglonGrid + 1
                    grGrafico.Row = lRenglonGrid
                    
                    'grGrafico.Row = grGrafico.RowCount
                    grGrafico.RowLabel = oCampo.Valor
                    
                    seriePeriodo = oCampo.Valor
                Else
                    
                    grGrafico.Row = lRenglon
                    
                End If


            End If

            Set oCampo = Registro(3)
            lColumna = yaExisteEtiquetaColumna(CStr(oCampo.Valor), cEtiquetaColumna)
            
            If lColumna = 0 Then ' NO existe columna, creala y verifica llenar los espacios hacia arriba, para completar la serie
                
                cEtiquetaColumna.Add CStr(oCampo.Valor)
                
                If cEtiquetaColumna.Count > grGrafico.ColumnCount Then
                    grGrafico.ColumnCount = grGrafico.ColumnCount + 1
                End If
                
                lColumnaGrid = lColumnaGrid + 1
                grGrafico.Column = lColumnaGrid
                
                grGrafico.ColumnLabel = oCampo.Valor
                
                If grGrafico.Row > 1 Then
                    lRenglonTemp = grGrafico.Row
                    For lRenglon = 1 To lRenglonTemp
                        
                        grGrafico.Row = lRenglon
                        If lRenglonTemp = lRenglon Then
                            Set oCampo = Registro(2)
                            If IsNull(oCampo.Valor) Then
                                grGrafico.Data = 0
                            Else
                                grGrafico.Data = oCampo.Valor
                            End If
                        Else
                            grGrafico.Data = 0
                        End If
                    
                    Next lRenglon
                
                Else
                
                    Set oCampo = Registro(2)
                    
                    If IsNull(oCampo.Valor) Then
                        grGrafico.Data = 0
                    Else
                        grGrafico.Data = oCampo.Valor
                    End If
                
                End If
                
            Else
            
                grGrafico.Column = lColumna
                
                Set oCampo = Registro(2)
                
                If IsNull(oCampo.Valor) Then
                    grGrafico.Data = 0
                Else
                    grGrafico.Data = oCampo.Valor
                End If
                
            End If

        Next Registro
        
        grGrafico.ColumnCount = lColumnaGrid
        grGrafico.RowCount = lRenglonGrid
        
        dibujaGrafica = True
    
    Else

        dibujaGrafica = False

    End If
    
End Function

Private Function yaExisteEtiquetaColumna(strValor As String, cSeries As Collection) As Long
        
    Dim i As Long
    
    For i = 1 To cSeries.Count
    
        If strValor = cSeries.Item(i) Then
            yaExisteEtiquetaColumna = i
            Exit For
        End If
    
    Next i
        
End Function

Private Function yaExisteEtiquetaRenglon(strValor As String, cSeries As Collection) As Long
        
    Dim i As Long
    
    For i = 1 To cSeries.Count
    
        If strValor = cSeries.Item(i) Then
            yaExisteEtiquetaRenglon = i
            Exit For
        End If
    
    Next i
        
End Function

