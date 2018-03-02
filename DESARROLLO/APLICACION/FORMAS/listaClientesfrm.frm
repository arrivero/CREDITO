VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form listaClientesfrm 
   Caption         =   "Resultado (Posibles Clientes)"
   ClientHeight    =   3225
   ClientLeft      =   2415
   ClientTop       =   3120
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   7245
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2475
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7185
      _Version        =   196608
      _ExtentX        =   12674
      _ExtentY        =   4366
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      SpreadDesigner  =   "listaClientesfrm.frx":0000
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Obtener Dato"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "listaClientesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer


Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub cmdsalir_Click()

Dim datos As Recordset

grdpacientes.Col = 1
clave = grdpacientes.Text

If clave > 0 Then

    Set datos = base.OpenRecordset("select * from clientes where no_cliente=" & CStr(clave))
    If datos.RecordCount > 0 Then
    Select Case band
    Case 1
        frmclientes.txtnombre.Text = datos!Nombre
        frmclientes.txtapellido.Text = datos!apellido
        frmclientes.txtdireccion.Text = IIf(datos!DIRECCION = "Nulo", "", datos!DIRECCION)
        frmclientes.txtubicacion.Text = IIf(datos!ubicacion = "Nulo", "", datos!ubicacion)
        frmclientes.txtcd.Text = IIf(datos!ciudad = "Nulo", "", datos!ciudad)
        frmclientes.txtedo.Text = IIf(datos!estado = "Nulo", "", datos!estado)
        frmclientes.txtcp.Text = IIf(datos!cp = "Nulo", "", datos!cp)
        frmclientes.txttelefono.Text = IIf(datos!TELEFONO = "Nulo", "", datos!TELEFONO)
        frmclientes.txtmaxcredito.Text = datos!maxcredito
        limite = datos!maxcredito
        frmclientes.txtnocliente.Text = datos!no_cliente
        frmclientes.cmdcredito.Enabled = True
        frmclientes.cmdpagos.Enabled = True
        frmclientes.cmdnuevo.Enabled = True
        frmclientes.cmdgrabar.Enabled = True
        frmclientes.cmdgrabar.Caption = "Modificar"
        nocliente = datos!no_cliente
    Case 2
        frmconsultanombre.txtnombre.Text = datos!Nombre
        frmconsultanombre.txtapellido.Text = datos!apellido
        frmconsultanombre.txtnocliente = datos!no_cliente
        nocliente = datos!no_cliente
    Case 3
        frmconsultanombre.txtnombre.Text = datos!Nombre
        frmconsultanombre.txtapellido.Text = datos!apellido
        frmconsultanombre.txtnocliente = datos!no_cliente
        nocliente = datos!no_cliente
    End Select
    Else
        'cmdgrabar.Caption = "Dar de Alta"
    End If
    datos.Close
End If

Unload Me
End Sub

Private Sub Form_Load()

Dim datos As Recordset

i = 0

grdpacientes.Row = 0
grdpacientes.Col = 1
grdpacientes.Text = "No. Cliente"

grdpacientes.Col = 2
grdpacientes.Text = "Nombre"

grdpacientes.Col = 3
grdpacientes.Text = "Apellido"


If Nombre <> "" Then
    Set datos = base.OpenRecordset("select * from clientes where Ucase(nombre) like '" & UCase(Nombre) & "*'")
    While Not datos.EOF
        grdpacientes.Rows = grdpacientes.Rows + 1
        grdpacientes.Row = grdpacientes.Rows - 2
        
        grdpacientes.Col = 1
        grdpacientes.Text = datos!no_cliente
        
        grdpacientes.Col = 2
        grdpacientes.Text = datos!Nombre
        
        grdpacientes.Col = 3
        grdpacientes.Text = datos!apellido
        
        i = i + 1
        
        datos.MoveNext
    Wend
    If i > 0 Then
        grdpacientes.Col = 1
        grdpacientes.Row = 1
    Else
        clave = 0
        MsgBox "No existen clientes con ese nombre", vbInformation, "Consulta de Clientes"
    End If
End If
datos.Close
End Sub

