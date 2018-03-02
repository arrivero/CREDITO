VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form empleadoCambio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualización de Datos del Empleado"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Usuario:"
      Height          =   1005
      Left            =   45
      TabIndex        =   38
      Top             =   4770
      Width           =   6225
      Begin VB.CheckBox chkEstatusUsuario 
         Caption         =   "Status Usuario Activo"
         Height          =   285
         Left            =   4140
         TabIndex        =   43
         Top             =   540
         Width           =   1950
      End
      Begin VB.ComboBox cbPerfil 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2070
         TabIndex        =   41
         Top             =   540
         Width           =   1860
      End
      Begin VB.TextBox txtUsuario 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   90
         MaxLength       =   20
         TabIndex        =   39
         Top             =   540
         Width           =   1860
      End
      Begin VB.Label lblPerfil 
         Caption         =   "Perfil:"
         Height          =   255
         Left            =   2070
         TabIndex        =   42
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "ID Usuario:"
         Height          =   255
         Left            =   90
         TabIndex        =   40
         Top             =   270
         Width           =   1410
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Direcci{on:"
      Height          =   1680
      Left            =   0
      TabIndex        =   27
      Top             =   3060
      Width           =   6270
      Begin VB.TextBox txtCalleNumero 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   90
         MaxLength       =   20
         TabIndex        =   32
         Top             =   525
         Width           =   3975
      End
      Begin VB.TextBox txtColonia 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4185
         MaxLength       =   20
         TabIndex        =   31
         Top             =   540
         Width           =   1860
      End
      Begin VB.TextBox txtMunicipio 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   90
         MaxLength       =   20
         TabIndex        =   30
         Top             =   1170
         Width           =   1860
      End
      Begin VB.TextBox txtEstado 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2070
         MaxLength       =   20
         TabIndex        =   29
         Top             =   1170
         Width           =   1860
      End
      Begin VB.TextBox txtCodigoPostal 
         BackColor       =   &H8000000F&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4185
         MaxLength       =   20
         TabIndex        =   28
         Top             =   1170
         Width           =   1860
      End
      Begin VB.Label Label1 
         Caption         =   "Calle - Número:"
         Height          =   255
         Left            =   135
         TabIndex        =   37
         Top             =   270
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "Colonia:"
         Height          =   255
         Left            =   4185
         TabIndex        =   36
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label3 
         Caption         =   "Municipio:"
         Height          =   255
         Left            =   90
         TabIndex        =   35
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label Label4 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   2070
         TabIndex        =   34
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "C.P.:"
         Height          =   255
         Left            =   4185
         TabIndex        =   33
         Top             =   900
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Genrales:"
      Height          =   3030
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6270
      Begin VB.Frame Frame2 
         Caption         =   "Género:"
         Height          =   645
         Left            =   2160
         TabIndex        =   14
         Top             =   945
         Width           =   1860
         Begin VB.OptionButton opGenero 
            Caption         =   "Hombre"
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   16
            Top             =   225
            Width           =   915
         End
         Begin VB.OptionButton opGenero 
            Caption         =   "Mujer"
            Height          =   285
            Index           =   2
            Left            =   990
            TabIndex        =   15
            Top             =   225
            Width           =   690
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   90
         MaxLength       =   20
         TabIndex        =   13
         Top             =   495
         Width           =   1860
      End
      Begin VB.TextBox txtApellidoPaterno 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   12
         Top             =   495
         Width           =   1860
      End
      Begin VB.TextBox txtApellidoMaterno 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4230
         MaxLength       =   20
         TabIndex        =   11
         Top             =   495
         Width           =   1860
      End
      Begin VB.ComboBox cbEstadoCivil 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   4185
         TabIndex        =   10
         Top             =   1260
         Width           =   1860
      End
      Begin VB.TextBox txtRFC 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   135
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2025
         Width           =   1860
      End
      Begin VB.TextBox txtCURP 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2115
         MaxLength       =   20
         TabIndex        =   8
         Top             =   2025
         Width           =   1860
      End
      Begin VB.TextBox txtTelefono 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   135
         MaxLength       =   15
         TabIndex        =   7
         Top             =   2700
         Width           =   1860
      End
      Begin VB.TextBox txtCorreo 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2115
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2700
         Width           =   1860
      End
      Begin SSCalendarWidgets_A.SSDateCombo dFechaNacimiento 
         Height          =   330
         Left            =   135
         TabIndex        =   17
         Top             =   1260
         Width           =   1860
         _Version        =   65537
         _ExtentX        =   3281
         _ExtentY        =   582
         _StockProps     =   93
         BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinDate         =   "1960/1/1"
         ShowCentury     =   -1  'True
      End
      Begin VB.Label Label13 
         Caption         =   "Fec Nac.:"
         Height          =   255
         Left            =   135
         TabIndex        =   26
         Top             =   945
         Width           =   1905
      End
      Begin VB.Label Label17 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   135
         TabIndex        =   25
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label16 
         Caption         =   "Ap. Paterno:"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label15 
         Caption         =   "Ap. Materno:"
         Height          =   255
         Left            =   4230
         TabIndex        =   23
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Estado Civil:"
         Height          =   255
         Left            =   4185
         TabIndex        =   22
         Top             =   945
         Width           =   945
      End
      Begin VB.Label Label20 
         Caption         =   "RFC:"
         Height          =   255
         Left            =   135
         TabIndex        =   21
         Top             =   1710
         Width           =   1410
      End
      Begin VB.Label Label14 
         Caption         =   "CURP:"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label22 
         Caption         =   "Teléfono:"
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label Label21 
         Caption         =   "e-Mail:"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   2385
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza"
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6570
      Width           =   1215
   End
   Begin VB.ComboBox cbEstadoEmpleado 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   6165
      Width           =   1860
   End
   Begin VB.TextBox txtSalario 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2058
         SubFormatType   =   2
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4140
      MaxLength       =   20
      TabIndex        =   0
      Top             =   6165
      Width           =   2085
   End
   Begin VB.Label Label19 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   5850
      Width           =   630
   End
   Begin VB.Label Label23 
      Caption         =   "Salario:"
      Height          =   255
      Left            =   4185
      TabIndex        =   3
      Top             =   5850
      Width           =   945
   End
End
Attribute VB_Name = "empleadoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvariEmpleado As Integer
Private mbAlta As Boolean

Public Property Let bAlta(ByVal vData As Boolean)
    mbAlta = vData
End Property

Public Property Get bAlta() As Boolean
    bAlta = mbAlta
End Property

Public Property Let iEmpleado(ByVal vData As Integer)
    mvariEmpleado = vData
End Property

Public Property Get iEmpleado() As Integer
    iEmpleado = mvariEmpleado
End Property

Private Sub Form_Load()
    
    If mvariEmpleado > 0 Then
    
        Dim oEmpleado As New Empleado
        
         If oEmpleado.Informacion(mvariEmpleado) = True Then
         
            If oEmpleado.catalogoEstado() = True Then
                fnLlenaComboCollecion cbEstadoEmpleado, oEmpleado.cDatos, 0, ""
                cbEstadoEmpleado.ListIndex = -1
            End If
            
            If oEmpleado.catalogoEstadoCivil() = True Then
                fnLlenaComboCollecion cbEstadoCivil, oEmpleado.cDatos, 0, ""
                cbEstadoCivil.ListIndex = -1
            End If
            
            If oEmpleado.catalogoPerfil() = True Then
                fnLlenaComboCollecion cbPerfil, oEmpleado.cDatos, 0, ""
                cbPerfil.ListIndex = -1
            End If
                
            Me.txtNombre.Text = oEmpleado.strNombre
            Me.txtApellidoPaterno = oEmpleado.strApPaterno
            Me.txtApellidoMaterno = oEmpleado.strApMaterno
            Me.dFechaNacimiento.Text = oEmpleado.fNacimiento
            If oEmpleado.iGenero = 1 Then
                opGenero.Item(1).Value = True
            End If
            If oEmpleado.iGenero = 2 Then
                opGenero.Item(2).Value = True
            End If
            Call fnBuscaElemento(cbEstadoCivil, oEmpleado.iEstadoCivil)
            Me.txtRFC = oEmpleado.strRFC
            Me.txtCURP = oEmpleado.strCURP
            Me.txtTelefono = oEmpleado.strTelefono
            Me.txtCorreo = oEmpleado.strCorreo
            Me.txtCalleNumero = oEmpleado.strDireccion
            Me.txtColonia = oEmpleado.strColonia
            Me.txtMunicipio = oEmpleado.strCiudad
            Me.txtEstado = oEmpleado.strEstado
            Me.txtCodigoPostal = oEmpleado.strCP
            Me.txtUsuario = oEmpleado.strUsuario
            Call fnBuscaElemento(cbPerfil, oEmpleado.iPerfil)
            Me.chkEstatusUsuario.Value = oEmpleado.iEstadoUsuario
            
         End If
        
        Set oEmpleado = Nothing
        
    End If
    
End Sub

Private Sub cmdActualiza_Click()

    If validaDatosEmpleado() = True Then
        
        Dim cDatos As New Collection
        Dim Registro As New Collection
        Dim oCampo As New Campo
        Dim iGenero As Integer
        
        If opGenero.Item(1).Value = True Then
            iGenero = 1
        End If
        If opGenero.Item(2).Value = True Then
            iGenero = 2
        End If
    
        Registro.Add oCampo.CreaCampo(adInteger, , , mvariEmpleado)
        Registro.Add oCampo.CreaCampo(adInteger, , , cbEstadoEmpleado.ItemData(cbEstadoEmpleado.ListIndex))
        Registro.Add oCampo.CreaCampo(adInteger, , , txtNombre.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtApellidoPaterno.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtApellidoMaterno.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.dFechaNacimiento.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtRFC.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtCURP.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , cbEstadoCivil.ItemData(cbEstadoCivil.ListIndex))
        Registro.Add oCampo.CreaCampo(adInteger, , , iGenero)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtTelefono.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtCorreo.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Val(fnstrValor(Me.txtSalario.Text)))
    
        Registro.Add oCampo.CreaCampo(adInteger, , , txtCalleNumero.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , txtColonia.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtMunicipio.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtEstado.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , Me.txtCodigoPostal.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , cbPerfil.ItemData(cbPerfil.ListIndex))
        If Me.chkEstatusUsuario.Value = "1" Then
            Registro.Add oCampo.CreaCampo(adInteger, , , 1)
        Else
            Registro.Add oCampo.CreaCampo(adInteger, , , 0)
        End If
        
        cDatos.Add Registro
            
        Dim oEmpleado As New Empleado
                        
            Call oEmpleado.actualiza(cDatos)
            
        Set oEmpleado = Nothing
                
        mbAlta = True
        
        Unload Me
        
    End If

End Sub

Private Function validaDatosEmpleado() As Boolean

    validaDatosEmpleado = False
    
    If Len(txtNombre.Text) <= 0 Then
        MsgBox "Ingrese el nombre del Empleado", vbOKOnly + vbInformation
        txtNombre.SetFocus
        Exit Function
    End If

    If Len(txtApellidoPaterno.Text) <= 0 Then
        MsgBox "Ingrese el Apellido paterno del Empleado", vbOKOnly + vbInformation
        txtApellidoPaterno.SetFocus
        Exit Function
    End If

    If Len(dFechaNacimiento.Text) <= 0 Then
        MsgBox "Ingrese la fecha de nacimiento del Empleado", vbOKOnly + vbInformation
        dFechaNacimiento.SetFocus
        Exit Function
    End If

    If opGenero.Item(1).Value = False And opGenero.Item(2).Value = False Then
        MsgBox "Defina el género del Empleado", vbOKOnly + vbInformation
        opGenero.Item(1).SetFocus
        Exit Function
    End If

    If cbEstadoCivil.ListIndex < 0 Then
        MsgBox "Defina el Estado Civil del Empleado", vbOKOnly + vbInformation
        cbEstadoCivil.SetFocus
        Exit Function
    End If

    If cbEstadoEmpleado.ListIndex < 0 Then
        MsgBox "Defina el Estado actual del Empleado", vbOKOnly + vbInformation
        cbEstadoEmpleado.SetFocus
        Exit Function
    End If

    If Len(txtCURP.Text) <= 0 Then
        MsgBox "Ingrese el CURP del Empleado", vbOKOnly + vbInformation
        txtCURP.SetFocus
        Exit Function
    End If

    If Len(txtTelefono.Text) <= 0 Then
        MsgBox "Ingrese el Número de Teléfono del Empleado", vbOKOnly + vbInformation
        txtTelefono.SetFocus
        Exit Function
    End If

'DATOS DE DIRECCIÓN
    If Len(txtCalleNumero.Text) <= 0 Then
        MsgBox "Ingrese la calle y el número de la dirección del Empleado", vbOKOnly + vbInformation
        txtCalleNumero.SetFocus
        Exit Function
    End If

    If Len(txtColonia.Text) <= 0 Then
        MsgBox "Ingrese la colonia de la dirección del Empleado", vbOKOnly + vbInformation
        txtColonia.SetFocus
        Exit Function
    End If

    If Len(txtMunicipio.Text) <= 0 Then
        MsgBox "Ingrese el municipio de la dirección del Empleado", vbOKOnly + vbInformation
        txtMunicipio.SetFocus
        Exit Function
    End If

    If Len(txtEstado.Text) <= 0 Then
        MsgBox "Ingrese el Estado de la dirección del Empleado", vbOKOnly + vbInformation
        txtEstado.SetFocus
        Exit Function
    End If

    If Len(txtCodigoPostal.Text) <= 0 Then
        MsgBox "Ingrese el Código Postal de la dirección del Empleado", vbOKOnly + vbInformation
        txtCodigoPostal.SetFocus
        Exit Function
    End If

'DATOS DE SEGURIDAD

    If Len(txtUsuario.Text) <= 0 Then
        MsgBox "Ingrese el Usuario para acceso a la aplicación para el Empleado", vbOKOnly + vbInformation
        txtUsuario.SetFocus
        Exit Function
    End If

    validaDatosEmpleado = True
    
End Function

