VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form sicPrincipalfrm 
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11070
   Icon            =   "sicPrincipalfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   10170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   17939
      _Version        =   196608
      AutoSize        =   1
      SplitterBarWidth=   3
      BorderStyle     =   3
      BackColor       =   16777215
      Locked          =   -1  'True
      PaneTree        =   "sicPrincipalfrm.frx":2372
      Begin Threed.SSPanel SSPanel1 
         Height          =   9375
         Left            =   30
         TabIndex        =   1
         Top             =   765
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   16536
         _Version        =   196608
         BackColor       =   16777215
         PictureAlignment=   7
         RoundedCorners  =   0   'False
         FloodShowPct    =   0   'False
         Begin MSComDlg.CommonDialog cmdEstadosCuenta 
            Left            =   1590
            Top             =   9240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            DialogTitle     =   "Genera Pagos Internet"
            Flags           =   1
            FontSize        =   10
         End
         Begin Crystal.CrystalReport crCorteDiario 
            Left            =   1620
            Top             =   7290
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin Crystal.CrystalReport crReporteHandHeld 
            Left            =   1620
            Top             =   8610
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            PrintFileName   =   "C:\Archivos INTEGRA\rephh"
            PrintFileType   =   2
            PrintFileLinesPerPage=   60
         End
         Begin Threed.SSRibbon cmdRepHandHeld 
            Height          =   645
            Left            =   0
            TabIndex        =   2
            Top             =   8520
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":2404
            Caption         =   "Reporte Hand Held"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdUsuarioPagos 
            Height          =   645
            Left            =   0
            TabIndex        =   3
            Top             =   1320
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":4BB6
            Caption         =   "Usuarios - Pagos"
            PictureAlignment=   0
         End
         Begin Threed.SSRibbon cmdCaja 
            Height          =   645
            Left            =   0
            TabIndex        =   4
            Top             =   1950
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            BackColor       =   12648447
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":6F38
            Caption         =   "Movimientos Caja"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdCreditos 
            Height          =   645
            Left            =   0
            TabIndex        =   5
            Top             =   5400
            Visible         =   0   'False
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            BackColor       =   16777215
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":738A
            Caption         =   "Créditos"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdpagos 
            Height          =   645
            Left            =   0
            TabIndex        =   6
            Top             =   660
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            BackColor       =   -2147483645
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":76A4
            Caption         =   "Pagos"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdClientes 
            Height          =   645
            Left            =   0
            TabIndex        =   7
            Top             =   30
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            CaptionStyle    =   1
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":7AF6
            Caption         =   "Clientes"
            PictureAlignment=   1
            PictureDnChange =   2
         End
         Begin Threed.SSRibbon cmdReporteGeneral 
            Height          =   645
            Left            =   0
            TabIndex        =   8
            Top             =   7890
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":7F48
            Caption         =   "Reporte General"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdConfiguracion 
            Height          =   645
            Left            =   0
            TabIndex        =   9
            Top             =   2580
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            BackColor       =   12640511
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":A6FA
            Caption         =   "Configuración"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdCorteDiario 
            Height          =   645
            Left            =   0
            TabIndex        =   10
            Top             =   7260
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":AB4C
            Caption         =   "Corte Diario"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdEstadistico 
            Height          =   645
            Left            =   0
            TabIndex        =   11
            Top             =   6630
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":B616
            Caption         =   "Resumen"
            PictureAlignment=   1
         End
         Begin Threed.SSRibbon cmdReporteInternet 
            Height          =   645
            Left            =   0
            TabIndex        =   12
            Top             =   9150
            Visible         =   0   'False
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   1138
            _Version        =   196608
            PictureFrames   =   1
            Picture         =   "sicPrincipalfrm.frx":BA68
            Caption         =   "Reporte - Internet"
            PictureAlignment=   0
         End
      End
      Begin Threed.SSPanel pnlTitulo 
         Height          =   690
         Left            =   2175
         TabIndex        =   13
         Top             =   30
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   1217
         _Version        =   196608
         ForeColor       =   -2147483641
         BackColor       =   -2147483646
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   0   'False
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   675
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1191
         _Version        =   196608
         BackColor       =   -2147483646
         RoundedCorners  =   0   'False
         FloodShowPct    =   0   'False
      End
   End
End
Attribute VB_Name = "sicPrincipalfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEstadistico_Click(Value As Integer)
    
    If Value = 0 Then
        Me.Caption = ""
        pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
        despliegaVentana portadafrm, WND_PORTADA
        Exit Sub
    Else
        If iVentana = WND_RESUMEN_ANALISIS Then
            Exit Sub
        End If
    End If
    
    accesofrm.Show vbModal

    If giTipoUsuario <> USUARIO_GERENTE Then

        Exit Sub
        
    End If

    If accesofrm.bPermiteAcceso = True Then
    
        Me.Caption = "Solución Integral de Administración de Creditos"
        pnlTitulo.Caption = "Acumulado - Porcentajes"
        
        despliegaVentana resumenGraficofrm, WND_RESUMEN_ANALISIS
    
    End If

End Sub

Private Sub cmdReporteInternet_Click(Value As Integer)

    If Value = 0 Then
        Exit Sub
    Else

        Dim strArchivo As String
        'Dim iNumeroArchivo As Integer
    
        cmdEstadosCuenta.Filter = "qryInternet(*.txt)|*.txt"
        cmdEstadosCuenta.FileName = "qryInternet"
        cmdEstadosCuenta.DialogTitle = "Estados de Cuenta."
        cmdEstadosCuenta.ShowOpen
        
        If Len(cmdEstadosCuenta.FileName) > 12 Then
        
            strArchivo = cmdEstadosCuenta.FileName
        
            'Poner el cursor de proceso
            Screen.MousePointer = vbHourglass
            
            'Hacer la consulta de los pagos de los créditos vigentes
            Dim oCredito As New credito
            Dim strMensaje As String
            'Abrir archivo para escritura
            'If abreArchivofn(strArchivo, iNumeroArchivo, PARA_ESCRITURA) = True Then
            
                'Obten los pagos (estados de cuenta)
                strMensaje = oCredito.obtenPagosInternet(strArchivo)
                
                'If oCredito.cDatos.Count > 0 Then
                    
                '    'Guardar los pagos en el archivo
                '    agregaRegistrosArchivoInternet iNumeroArchivo, oCredito.cDatos
                    
                '    If oCredito.cDatos.Count > 0 Then
                '        For i = 1 To oCredito.cDatos.Count
                '            oCredito.cDatos.Remove 1
                '        Next
                '    End If

                '    Set oCredito = Nothing
                    
                    'Cerrar el archivo
                '    cierraArchivofn iNumeroArchivo
                    
                    'Enviar mensaje de proceso terminado
                    MsgBox strMensaje, vbInformation + vbOKOnly
                    
                    'cmdReporteInternet.SetFocus
                    
                'Else
                '    MsgBox "Los estados de cuenta no se generaron, por favor intente de nuevo", vbInformation + vbOKOnly
                'End If
                
            'Else
            '    MsgBox "Los estados de cuenta no se generaron, por favor intente de nuevo", vbInformation + vbOKOnly
            'End If
            
            
        Else
            Exit Sub
        End If
        
        'Poner el cursor Normal
        Screen.MousePointer = vbDefault
    
    End If
    
End Sub

Private Sub Form_Load()
    
    bInicio = True
    
    Me.Caption = ""
    pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
    despliegaVentana portadafrm, WND_PORTADA
        
    Call fnObtenConfiguracionIni
    
    'Actualiza creditos, descuentos, atrasos y dias de atraso
    If fnProcedeActualizarCreditos = True Then
    
        Dim oCredito As New credito
        Dim strFecha As String
        
        strFecha = Format(Now(), "dd/mm/yyyy")
        
        'Poner el cursor de proceso
        Screen.MousePointer = vbHourglass
        
        oCredito.actualizaCreditos strFecha
                
        SaveSetting NOMBRE_SOLUCION, "CREDITOS", "FECHA", Format(Now(), "dd/mm/yyyy")
                    
        Set oCredito = Nothing
        
        'Poner el cursor Normal
        Screen.MousePointer = vbDefault
        
    End If
        
End Sub


Private Function fnProcedeActualizarCreditos() As Boolean
    
    Dim strFechaUltima As String
    
    strFechaUltima = GetSetting(NOMBRE_SOLUCION, "CREDITOS", "FECHA", "0")
    If strFechaUltima = "0" Then
        SaveSetting NOMBRE_SOLUCION, "CREDITOS", "FECHA", Format(Now(), "dd/mm/yyyy")
        fnProcedeActualizarCreditos = True
    Else
        If DateDiff("d", strFechaUltima, Format(Now(), "dd/mm/yyyy")) = 0 Then
            fnProcedeActualizarCreditos = False
        Else
            fnProcedeActualizarCreditos = True
        End If
        
    End If

End Function

Private Sub Form_Unload(Cancel As Integer)
    
    Unload oFormaActual
    
End Sub

Private Sub cmdClientes_Click(Value As Integer)
    
    If Value = 0 Then
        Me.Caption = ""
        pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
        despliegaVentana portadafrm, WND_PORTADA
        Exit Sub
    Else
        If iVentana = WND_CLIENTES Then
            Exit Sub
        End If
    End If
    
    accesofrm.Show vbModal

    If accesofrm.bPermiteAcceso = True Then
        
        Me.Caption = "Solución Integral de Administración de Creditos"
        pnlTitulo.Caption = "Clientes - (" & UCase(gstrUsuario) & ")"
        
        despliegaVentana clientesfrm, WND_CLIENTES
        
    End If
    
End Sub


Private Sub cmdCreditos_Click(Value As Integer)
    
    If Value = 0 Then
        Me.Caption = ""
        pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
        despliegaVentana portadafrm, WND_PORTADA
        Exit Sub
    Else
        If iVentana = WND_CREDITOS Then
            Exit Sub
        End If
    End If
        
    accesofrm.Show vbModal

    If accesofrm.bPermiteAcceso = True Then
    
        Me.Caption = "Solución Integral de Administración de Creditos"
        pnlTitulo.Caption = "Créditos - (" & UCase(gstrUsuario) & ")"
        
        despliegaVentana creditosfrm, WND_CREDITOS
        
    End If
    
End Sub

Private Sub cmdpagos_Click(Value As Integer)
    
    If Value = 0 Then
        Me.Caption = ""
        pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
        despliegaVentana portadafrm, WND_PORTADA
        Exit Sub
    Else
        If iVentana = WND_PAGOS Then
            Exit Sub
        ElseIf iVentana = WND_PAGOS_NUEVOS Then
            Exit Sub
        End If
    End If
    
    accesofrm.Show vbModal
    
    If accesofrm.bPermiteAcceso = True Then
        Dim lResuldato As Long
        Dim strUsuario As String
        strUsuario = gstrUsuario
    
        Me.Caption = "Solución Integral de Administración de Creditos"
        pnlTitulo.Caption = "Pagos - (" & UCase(gstrUsuario) & ")"

        automaticoManualfrm.Show vbModal

        If automaticoManualfrm.iAutomaticoManual = 1 Then

            'If vbCancel = MsgBox("Coloque su equipo en la base y descargue los pagos, al terminar haga clic en 'Ok' para registrar sus pagos", vbInformation + vbOKCancel) Then
            '    Exit Sub
            'Else

                If gEncriptado = "NO" Then
                
                    pagoNuevofrm.iAutomaticoManual = 1
                    pagoNuevofrm.strUsuario = strUsuario
                    despliegaVentana pagoNuevofrm, WND_PAGOS_NUEVOS
                    
                Else
                    
                    Dim strNombreUsuario As String
                    
                    usuarioPagos.Show vbModal
                    strNombreUsuario = usuarioPagos.strNombreUsuario
                    
                    If Len(strNombreUsuario) > 0 Then
                        
                        lResuldato = ShellExecute(Me.hwnd, "", App.Path & "\Decrypt\Decrypt.exe", "", "", 3)
                        
                        If lResuldato = 42 Then
                        
                            MsgBox "Ahora se mostraran los pagos.", vbOKOnly
                            'MsgBox "Antes de continuar, primero seleccione el achivo a integrar", vbOKOnly
                            
                            pagoNuevofrm.iAutomaticoManual = 1
                            pagoNuevofrm.strUsuario = strUsuario
                            pagoNuevofrm.strNombreUsuarioPagos = strNombreUsuario
                            despliegaVentana pagoNuevofrm, WND_PAGOS_NUEVOS
                            
                        End If
                    
                    End If
                    
                End If
                
            'End If

        Else

            capturaPagofrm.nombreaux = strUsuario
            
            despliegaVentana capturaPagofrm, WND_PAGOS
        
        End If

    End If

End Sub

Private Sub cmdCaja_Click(Value As Integer)
    
    If Value = 0 Then
        Me.Caption = ""
        pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
        despliegaVentana portadafrm, WND_PORTADA
        Exit Sub
    Else
        If iVentana = WND_MOVIMIENTOS Then
            Exit Sub
        End If
    End If
                
    accesofrm.Show vbModal

    If accesofrm.bPermiteAcceso = True Then
    
        Dim strUsuario As String
        strUsuario = gstrUsuario

        Me.Caption = "Solución Integral de Administración de Creditos"
        If giTipoUsuario = USUARIO_GERENTE Then
            pnlTitulo.Caption = "Movimientos de Caja - (" & UCase(gstrUsuario) & ")"
        Else
            pnlTitulo.Caption = "Gastos - (" & UCase(gstrUsuario) & ")"
        End If
        
        despliegaVentana gastosTrabajofrm, WND_MOVIMIENTOS
        
    End If
    
'    'bandm = 4
'    'gastosfrm.Show
'    pagosPercepcionesfrm.Show
'    SSSplitter1.Panes(2).Control = pagosPercepcionesfrm.hWnd
'    Set oFormaActual = pagosPercepcionesfrm
    
End Sub

Private Sub cmdConfiguracion_Click(Value As Integer)

    If Value = 0 Then
        Me.Caption = ""
        pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
        despliegaVentana portadafrm, WND_PORTADA
        Exit Sub
    Else
        If iVentana = WND_CONFIGURACION Then
            Exit Sub
        End If
    End If
    
    accesofrm.Show vbModal

    If accesofrm.bPermiteAcceso = True Then
    
        Me.Caption = "Solución Integral de Administración de Creditos"
        pnlTitulo.Caption = "Configuración - (" & UCase(gstrUsuario) & ")"
        
        despliegaVentana configuracionfrm, WND_CONFIGURACION
    
    End If

End Sub

Private Sub cmdUsuarioPagos_Click(Value As Integer)
    
    If Value = 0 Then
        Exit Sub
    End If
    
    repPagosUsuariofrm.Show vbModal

End Sub

Private Sub cmdRepHandHeld_Click(Value As Integer)

    If Value = 0 Then
        Exit Sub
    End If
        
    Dim oCredito As New credito
    Dim strFecha As String
    Dim strArchivo As String

    strFecha = DateAdd("d", 1, Format(Now(), "dd/mm/yyyy"))

    cmdEstadosCuenta.Filter = "rephh(*.chr)|*.chr"
    cmdEstadosCuenta.FileName = "rephh"
    cmdEstadosCuenta.DialogTitle = "Archivo Hand Held."
    cmdEstadosCuenta.ShowOpen
    
    If Len(cmdEstadosCuenta.FileName) > 12 Then
    
        strArchivo = cmdEstadosCuenta.FileName
    
        'Poner el cursor de proceso
        Screen.MousePointer = vbHourglass
        
        Dim strMensaje As String
        
        oCredito.actualizaCreditos strFecha
        
        'Obten los estados de cuenta, de los clientes en el reporte HH
        strMensaje = oCredito.obtenEstadosCuentaHH(strArchivo)
        
        'Poner el cursor Normal
        Screen.MousePointer = vbDefault
        
        'Enviar mensaje de proceso terminado
        MsgBox strMensaje, vbInformation + vbOKOnly
        
    Else
        Exit Sub
    End If
    
    Set oCredito = Nothing

'    Dim oReporte As New Reporte
'    Dim cParametros As New Collection
'    Dim oCampo As New Campo
'
'    cParametros.Add oCampo.CreaCampo(adInteger, , , Format(Now(), "dd/mm/yyyy"))
'    cParametros.Add oCampo.CreaCampo(adInteger, , , Format(Now(), "dd/mm/yyyy"))
'
'    oReporte.oCrystalReport = crReporteHandHeld
'    If gstrReporteEnPantalla = "Si" Then
'        oReporte.bVistaPreliminar = True
'    Else
'        oReporte.bVistaPreliminar = False
'    End If
'
'    oReporte.strImpresora = gPrintPed
'    oReporte.strNombreReporte = DirSys & "rephh.rpt"
'    oReporte.cParametros = cParametros
'    oReporte.fnImprime
'    Set oReporte = Nothing
'
'    'Poner el cursor Normal
'    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdCorteDiario_Click(Value As Integer)

    If Value = 0 Then
        Exit Sub
    End If

    'Poner el cursor de proceso
    Screen.MousePointer = vbHourglass

    Dim oCredito As New credito
    Dim strFecha As String
    
    oCredito.corteDiario
    Set oCredito = Nothing

    Dim oReporte As New Reporte
    Dim cParametros As New Collection
    Dim oCampo As New Campo
        
    cParametros.Add oCampo.CreaCampo(adInteger, , , Format(Now(), "dd/mm/yyyy"))
        
    oReporte.oCrystalReport = crCorteDiario
    If gstrReporteEnPantalla = "Si" Then
        oReporte.bVistaPreliminar = True
    Else
        oReporte.bVistaPreliminar = False
    End If
    oReporte.strImpresora = gPrintPed
    oReporte.strNombreReporte = DirSys & "rpCorteDiario.rpt"
    oReporte.cParametros = cParametros
    oReporte.fnImprime
    Set oReporte = Nothing

    'Poner el cursor Normal
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdReporteGeneral_Click(Value As Integer)
    
    If Value = 0 Then
        Exit Sub
    End If
    
    accesofrm.Show vbModal

    If accesofrm.bPermiteAcceso = True Then
    
        If giTipoUsuario = USUARIO_GERENTE Or giTipoUsuario = USUARIO_ADMINSTRADOR Then
        
            reGeneralfrm.Show vbModal
        
        End If
    
    End If
    
End Sub

