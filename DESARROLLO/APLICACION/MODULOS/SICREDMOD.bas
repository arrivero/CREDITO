Attribute VB_Name = "SICREDMOD"
Global Const MODIFICA_PAGO_SICRED_FUNCION = 0
Global Const MODIFICA_CREDITOS_SICRED_FUNCION = 1
Global Const REGISTRO_GASTOS_INTERNOS_SICRED_FUNCION = 2
Global Const REPORTE_OTROS_GASTOS_SICRED_FUNCION = 3
Global Const MODIFICACION_DE_DEPOSITOS_SICRED_FUNCION = 5
Global Const REGLAS_MORATORIOS_SICRED_FUNCION = 6
Global Const USUARIOS_SICRED_FUNCION = 7

Global Const ID_CONCEPTO_CAJA = 1
Global Const ST_CONCEPTO_CAJA = "CAJA"
Global Const ID_CONCEPTO_BANCOS = 2
Global Const ST_CONCEPTO_BANCOS = "BANCOS"
Global Const ID_CONCEPTO_INVENTARIO = 3
Global Const ST_CONCEPTO_INVENTARIO = "INVENTARIO"
Global Const ID_CONCEPTO_INGRESOS_POR_SERVICIOS = 4
Global Const ST_CONCEPTO_INGRESOS_POR_SERVICIOS = "Ingresos por servicios"
Global Const ID_CONCEPTO_INGRESOS_POR_PRODUCTOS = 5
Global Const ST_CONCEPTO_INGRESOS_POR_PRODUCTOS = "Ingresos por productos"
Global Const ID_CONCEPTO_INGRESOS_POR_TELEFONO_MONEDAS = 6
Global Const ST_CONCEPTO_INGRESOS_POR_TELEFONO_MONEDAS = "Ingresos por Teléfono Monedas"
Global Const ID_CONCEPTO_INGRESOS_POR_PUBLICIDAD = 7
Global Const ST_CONCEPTO_INGRESOS_POR_PUBLICIDAD = "Ingresos por publicidad PUBLICOM"
Global Const ID_CONCEPTO_EGRESOS_PAGO_RENTA = 8
Global Const ST_CONCEPTO_EGRESOS_PAGO_RENTA = "Pago de Renta"
Global Const ID_CONCEPTO_EGRESOS_PAGO_AGUA = 9
Global Const ST_CONCEPTO_EGRESOS_PAGO_AGUA = "Pago de Agua"
Global Const ID_CONCEPTO_EGRESOS_PAGO_LUZ = 10
Global Const ST_CONCEPTO_EGRESOS_PAGO_LUZ = "Pago de Luz"
Global Const ID_CONCEPTO_EGRESOS_PAGO_TELEFONO = 11
Global Const ST_CONCEPTO_EGRESOS_PAGO_TELEFONO = "Pago de Teléfono"
Global Const ID_CONCEPTO_EGRESOS_GASTOS_PAPELERIA = 12
Global Const ST_CONCEPTO_EGRESOS_GASTOS_PAPELERIA = "Gastos de Papelería"
Global Const ID_CONCEPTO_EGRESOS_GASTOS_PUBLICIDAD = 13
Global Const ST_CONCEPTO_EGRESOS_GASTOS_PUBLICIDAD = "Gastos de Publicidad"
Global Const ID_CONCEPTO_EGRESOS_GASTOS_PERIODICO = 14
Global Const ST_CONCEPTO_EGRESOS_GASTOS_PERIODICO = "Anuncio periódico-Contratación personal"
Global Const ID_CONCEPTO_EGRESOS_GASTOS_MATERIAL_LIMPIEZA = 15
Global Const ST_CONCEPTO_EGRESOS_GASTOS_MATERIAL_LIMPIEZA = "Compra de material para limpieza"
Global Const ID_CONCEPTO_EGRESOS_GASTOS_IMPRESION_PAPELERIA = 16
Global Const ST_CONCEPTO_EGRESOS_GASTOS_IMPRESION_PAPELERIA = "Impresión de papelería (formas)."
Global Const ID_CONCEPTO_EGRESOS_GASTOS_RADIOS_COMUNICACION = 17
Global Const ST_CONCEPTO_EGRESOS_GASTOS_RADIOS_COMUNICACION = "Radios de comunicación."
Global Const ID_CONCEPTO_EGRESOS_GASTOS_ASESORIA_CAPACITACION = 18
Global Const ST_CONCEPTO_EGRESOS_GASTOS_ASESORIA_CAPACITACION = "Asesoría y capacitación"
Global Const ID_CONCEPTO_EGRESOS_GASTOS_MANTENIMIENO_INSTALACION = 19
Global Const ST_CONCEPTO_EGRESOS_GASTOS_MANTENIMIENO_INSTALACION = "Mantenimiento instalaciones."
Global Const ID_CONCEPTO_EGRESOS_SALARIOS = 20
Global Const ST_CONCEPTO_EGRESOS_SALARIOS = "SALARIOS"
Global Const ID_CONCEPTO_EGRESOS_APORTACIONES_IMMS = 21
Global Const ST_CONCEPTO_EGRESOS_APORTACIONES_IMMS = "Aportaciones al IMMS."
Global Const ID_CONCEPTO_EGRESOS_APORTACIONES_INFONAVIT = 22
Global Const ST_CONCEPTO_EGRESOS_APORTACIONES_INFONAVIT = "Aportaciones al Infonavit."
Global Const ID_CONCEPTO_EGRESOS_SERVICIOS_CONTABLES = 23
Global Const ST_CONCEPTO_EGRESOS_SERVICIOS_CONTABLES = "Servicios contables."
Global Const ID_CONCEPTO_EGRESOS_SERVICIOS_LEGALES = 24
Global Const ST_CONCEPTO_EGRESOS_SERVICIOS_LEGALES = "Servicios legales."

'OPERACIONES
Global Const ID_OPERACION_INGRESOS_CLIENTES = 1
Global Const ST_OPERACION_INGRESOS_CLIENTES = "Ingresos Clientes."
Global Const ID_OPERACION_OTROS_INGRESOS = 2
Global Const ST_OPERACION_OTROS_INGRESOS = "Otros Ingresos."
Global Const ID_OPERACION_EGRESOS = 3
Global Const ST_OPERACION_EGRESOS = "Egresos."
Global Const ID_OPERACION_GASTOS_OPERACION = 4
Global Const ST_OPERACION_GASTOS_OPERACION = "Gastos de Operación."
Global Const ID_OPERACION_PAGO_PROVEEDORES = 5
Global Const ST_OPERACION_PAGO_PROVEEDORES = "Pago Proveedores."
Global Const ID_OPERACION_CAPITAL_SOCIAL = 6
Global Const ST_OPERACION_CAPITAL_SOCIAL = "Capital Social (Inversionistas)."

'ESTATUS DE CHEQUES
Global Const ST_CHEQUE_DISPONIBLE = 0
Global Const ST_CHEQUE_PAGADO = 1
Global Const ST_CHEQUE_COBRADO = 2
Global Const ST_CHEQUE_REBOTADO = 3
Global Const ST_CHEQUE_CANCELADO = 4

Global Const WND_PORTADA = 1
Global Const WND_CLIENTES = 2
Global Const WND_CREDITOS = 3
Global Const WND_PAGOS = 4
Global Const WND_PAGOS_NUEVOS = 5
Global Const WND_MOVIMIENTOS = 6
Global Const WND_CONFIGURACION = 7
Global Const WND_RESUMEN_ANALISIS = 8

Public lhwnd As Long
Public iVentana As Integer
Public oFormaActual As Object
Public bInicio As Boolean

Global Const SECCION_CLIENTE = 3

Public Function despliegaVentana(oForma As Object, iWnd As Integer)

    If bInicio = False Then
        Unload oFormaActual
    End If
    
    oForma.Show
    sicPrincipalfrm.SSSplitter1.Panes(SECCION_CLIENTE).Control = oForma.hWnd
    lhwnd = oForma.hWnd
    iVentana = iWnd
    Set oFormaActual = oForma

    bInicio = False
    
End Function

Public Function registraConceptos() As Boolean

    Dim cConceptos As New Collection
    Dim cConcepto As Collection
    Dim oCampo As New Campo
    
    Dim oCuenta As New cProveedor
    
    'ACTIVO CIRCULANTE
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_CAJA)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_CAJA)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_BANCOS)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_BANCOS)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_INVENTARIO)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_INVENTARIO)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_INGRESOS_POR_SERVICIOS)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_INGRESOS_POR_SERVICIOS)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_INGRESOS_POR_PRODUCTOS)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_INGRESOS_POR_PRODUCTOS)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_INGRESOS_POR_TELEFONO_MONEDAS)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_INGRESOS_POR_TELEFONO_MONEDAS)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_INGRESOS_POR_PUBLICIDAD)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_INGRESOS_POR_PUBLICIDAD)
    cConceptos.Add cConcepto
    
    'ACTIVO FIJO
    'cConcepto.Add oCampo.CreaCampo(adInteger, , , "Equipo de transporte")
    'cConcepto.Add oCampo.CreaCampo(adInteger, , , "Maquinaria y Equipo")
    'cConcepto.Add oCampo.CreaCampo(adInteger, , , "Préstamo a empleados")

    'PASIVO CIRCULANTE
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_RENTA)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_PAGO_RENTA)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_AGUA)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_PAGO_AGUA)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_LUZ)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_PAGO_LUZ)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_TELEFONO)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_PAGO_TELEFONO)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_PAPELERIA)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_PAPELERIA)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_PUBLICIDAD)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_PUBLICIDAD)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_PERIODICO)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_PERIODICO)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_MATERIAL_LIMPIEZA)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_MATERIAL_LIMPIEZA)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_IMPRESION_PAPELERIA)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_IMPRESION_PAPELERIA)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_RADIOS_COMUNICACION)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_RADIOS_COMUNICACION)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_ASESORIA_CAPACITACION)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_ASESORIA_CAPACITACION)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_MANTENIMIENO_INSTALACION)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_GASTOS_MANTENIMIENO_INSTALACION)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_SALARIOS)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_SALARIOS)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_APORTACIONES_IMMS)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_APORTACIONES_IMMS)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_APORTACIONES_INFONAVIT)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_APORTACIONES_INFONAVIT)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_SERVICIOS_CONTABLES)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_SERVICIOS_CONTABLES)
    cConceptos.Add cConcepto
    
    Set cConcepto = New Collection
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_SERVICIOS_LEGALES)
    cConcepto.Add oCampo.CreaCampo(adInteger, , , ST_CONCEPTO_EGRESOS_SERVICIOS_LEGALES)
    cConceptos.Add cConcepto
    
    Call oCuenta.registraConcepto(cConceptos)

    Set oCuenta = Nothing
    
End Function

Public Function registraTipoOperacion()
    
    Dim oCuenta As New cProveedor
    Dim oCampo As New Campo
    Dim cOperaciones As New Collection
    Dim cOperacion As Collection
    
    Set cOperacion = New Collection
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_INGRESOS_CLIENTES)
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ST_OPERACION_INGRESOS_CLIENTES)
    cOperaciones.Add cOperacion
    Set cOperacion = New Collection
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_OTROS_INGRESOS)
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ST_OPERACION_OTROS_INGRESOS)
    cOperaciones.Add cOperacion
    Set cOperacion = New Collection
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_EGRESOS)
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ST_OPERACION_EGRESOS)
    cOperaciones.Add cOperacion
    Set cOperacion = New Collection
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ST_OPERACION_GASTOS_OPERACION)
    cOperaciones.Add cOperacion
    Set cOperacion = New Collection
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_PAGO_PROVEEDORES)
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ST_OPERACION_PAGO_PROVEEDORES)
    cOperaciones.Add cOperacion
    Set cOperacion = New Collection
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_CAPITAL_SOCIAL)
    cOperacion.Add oCampo.CreaCampo(adInteger, , , ST_OPERACION_CAPITAL_SOCIAL)
    cOperaciones.Add cOperacion
    
    Call oCuenta.registraOperacion(cOperaciones)
    
    Set oCuenta = Nothing
    
End Function

Public Function registraOperacionCuenta(iSalon As Integer)
    
    Dim oCuenta As New cProveedor
    Dim oCampo As New Campo
    Dim cOperacionCuenta As Collection
    Dim cOperacionCuentas As New Collection
    
    'INGRESOS - ACTIVO CIRCULANTE
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon) 'iSALON
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_ACTIVO_CIRCULANTE_CAJA) 'CUENTA CONTABLE
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_INGRESOS_CLIENTES) 'oPERACION
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_CAJA) 'cONCEPTO
    cOperacionCuentas.Add cOperacionCuenta
    
    'ADICIONAL PARA REINVERSIÓN
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon) 'iSALON
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_ACTIVO_CIRCULANTE_CAJA) 'CUENTA CONTABLE
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_CAPITAL_SOCIAL) 'oPERACION
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_CAJA) 'cONCEPTO
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_ACTIVO_CIRCULANTE_BANCOS)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_INGRESOS_CLIENTES)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_BANCOS)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_ACTIVO_CIRCULANTE_ALMACEN)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_INGRESOS_CLIENTES)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_INVENTARIO)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_ACTIVO_CIRCULANTE_CLIENTES)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_INGRESOS_CLIENTES)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_INGRESOS_POR_SERVICIOS)
    cOperacionCuentas.Add cOperacionCuenta
    
    'EGRESOS - PACIVO CIRCULANTE
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_RENTA)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_RENTA)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_AGUA)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_AGUA)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_LUZ)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_LUZ)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_TELEFONO)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_PAGO_TELEFONO)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_PAPELERIA)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_PAPELERIA)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_PUBLICIDAD)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_PUBLICIDAD)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_ANUNCIO_PERIODICO)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_PERIODICO)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_MATS_LIMPIEZA)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_MATERIAL_LIMPIEZA)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_FORMAS_IMPRESAS)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_IMPRESION_PAPELERIA)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_RADIOS)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_RADIOS_COMUNICACION)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_ASESORIA_CAPACITACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_ASESORIA_CAPACITACION)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_MANTENIMIENTO)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_GASTOS_MANTENIMIENO_INSTALACION)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_SALARIOS)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_EGRESOS)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_SALARIOS)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_IMSS)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_EGRESOS)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_APORTACIONES_IMMS)
    cOperacionCuentas.Add cOperacionCuenta
    
    'EGRESOS - PACIVO CIRCULANTE - SALARIOS
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_INFONAVIT)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_APORTACIONES_INFONAVIT)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_SERVS_CONTABLES)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_SERVICIOS_CONTABLES)
    cOperacionCuentas.Add cOperacionCuenta
    
    Set cOperacionCuenta = New Collection
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , iSalon)
    'cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CUENTA_RESULTADOS_EGRESOS_SERVS_LEGALES)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_OPERACION_GASTOS_OPERACION)
    cOperacionCuenta.Add oCampo.CreaCampo(adInteger, , , ID_CONCEPTO_EGRESOS_SERVICIOS_LEGALES)
    cOperacionCuentas.Add cOperacionCuenta
    Call oCuenta.registraOperacionCuenta(cOperacionCuentas)
        
    Set oCuenta = Nothing
    
    'Dim oBD As New DataBase
    
    'Actualiza el parámetro modulo contable
    'oBD.Procedimiento = "CONFIGURACION_SALONUpdValorProc"
    'oBD.Parametros = iSalon & ", " & CONF_HAY_MODULO_CONTABLE & ", 'SI', " & 0
    'oBD.bFolio = False
    'oBD.fnInsert
    
    Set oBD = Nothing
            
End Function

