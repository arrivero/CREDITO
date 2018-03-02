VERSION 5.00
Begin VB.Form accesofrm 
   BorderStyle     =   0  'None
   Caption         =   "Bienvenido a tu Sistema de Control"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6285
   Icon            =   "accesofrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "-------------Bienvenido a tu Sistema de Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   30
      TabIndex        =   4
      Top             =   60
      Width           =   6195
      Begin VB.TextBox txtUsuario 
         Height          =   375
         Left            =   4410
         TabIndex        =   0
         Top             =   330
         Width           =   1695
      End
      Begin VB.TextBox txtClave 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4410
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   780
         Width           =   1695
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   645
         Left            =   3180
         Picture         =   "accesofrm.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1380
         Width           =   1425
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   645
         Left            =   4740
         Picture         =   "accesofrm.frx":0B04
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1380
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "USUARIO:"
         Height          =   255
         Left            =   3180
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "CLAVE:"
         Height          =   255
         Left            =   3180
         TabIndex        =   5
         Top             =   900
         Width           =   975
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2955
         Left            =   60
         Picture         =   "accesofrm.frx":163E
         Stretch         =   -1  'True
         Top             =   270
         Width           =   3075
      End
   End
End
Attribute VB_Name = "accesofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iFuncion As Integer 'Define que función se activa al autentificar el usuario.

Public bPermiteAcceso As Boolean

Private iVeces As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        cmdAceptar_Click
    
    End If
    
End Sub

Private Sub cmdAceptar_Click()

    bPermiteAcceso = False
    
    'Aplica criterios de seguridad de usuario
    Dim strMensaje As String
    
    strMensaje = fnValidaForma(Trim(txtClave.Text))
    
    'Call GeneraDigitoVerificador(Trim(txtUsuario.Text))
    
    If Len(strMensaje) > 0 Then
        
        MsgBox strMensaje, vbInformation, NOMBRE_SOLUCION + "-Inicio de sesión"
        txtClave.SetFocus
        txtClave.SelStart = 0
        txtClave.SelLength = Len(txtClave.Text)
        iVeces = iVeces + 1
        
    Else
        
        Dim bUsrPsw As Boolean
        Dim oUsuario As New Usuario
        If True = oUsuario.fnPermiteAcceso(Trim(txtUsuario.Text), Trim(txtClave.Text), bUsrPsw) Then
            
            gstrUsuario = Trim(txtUsuario.Text)
            giTipoUsuario = oUsuario.iTipo
            
            bPermiteAcceso = True
            
            Unload Me
            
            'Habrete sesamo
            'Select Case iFuncion 'bandm
            '    Case Is = MODIFICA_PAGO_SICRED_FUNCION
            '        frmmodificapago.Show vbModal
            '    Case Is = MODIFICA_CREDITOS_SICRED_FUNCION
            '        frmmcreditos.Show vbModal
            '    Case Is = REGISTRO_GASTOS_INTERNOS_SICRED_FUNCION
            '        frmgastosnos.Show vbModal
            '    Case Is = REPORTE_OTROS_GASTOS_SICRED_FUNCION
            '        frmotrosgastos.Show vbModal
            '    Case Is = 4
            '        accesa = 1
            '    Case Is = MODIFICACION_DE_DEPOSITOS_SICRED_FUNCION '5
            '        frmdepositos.Show vbModal
            '    Case Is = REGLAS_MORATORIOS_SICRED_FUNCION '6
            '        frmreglasmora.Show vbModal
            '    Case Is = USUARIOS_SICRED_FUNCION '7
            '        frmUsuarios.Show vbModal
            'End Select
            
        Else
            
            MsgBox "Verfique por favor su Usuario y/o Clave", vbInformation, NOMBRE_SOLUCION + "-Inicio de sesión"
            If bUsrPsw Then
                txtUsuario.SetFocus
                txtUsuario.SelStart = 0
                txtUsuario.SelLength = Len(txtUsuario.Text)
            Else
                txtClave.SetFocus
                txtClave.SelStart = 0
                txtClave.SelLength = Len(txtClave.Text)
            End If
            iVeces = iVeces + 1
            
        End If
        
        Set oUsuario = Nothing
        
    End If
    
    If iVeces = 3 Then
        Unload Me
    End If

End Sub

'Private Function controla() As Boolean
'
'    Dim strDia As String
'
'    controla = True
'
'    strDia = GetSetting(NOMBRE_SOLUCION, SECCION_APPCONFIG, LLAVE_DIAVALIDO)
'
'    If strDia = "" Then 'Inicia proceso de control
'
'        SaveSetting NOMBRE_SOLUCION, SECCION_APPCONFIG, LLAVE_DIAVALIDO, Format(Date, "dd/mm/yyyy")
'
'    Else
'        Dim iDias As Integer
'
'        iDias = DateDiff("d", CDate(Now), CDate(strDia))
'
'        If iDias >= 0 Then 'la fecha sistema es posterior a la fecha de control
'
'            iDias = DateDiff("d", CDate(Now), CDate(strDiaLimite))
'
'            If iDias < 0 Then 'se alcanzó limite de licencia, termina aplicación
'                controla = False
'            Else
'                SaveSetting NOMBRE_SOLUCION, SECCION_APPCONFIG, LLAVE_DIAVALIDO, Format(Date, "dd/mm/yyyy")
'            End If
'        Else    'hicieron cochupo, fin, adios, bye. Tranposos
'                controla = False
'        End If
'
'    End If
'
'End Function

Private Sub cmdcancelar_Click()
    bPermiteAcceso = False
    Unload Me
End Sub

Private Function fnValidaForma(strClave As String) As String

    fnValidaForma = ""
    'La longitud de la clave debe estar en un rango
    If Not (Len(strClave) >= 4 And Len(strClave) <= 10) Then
        fnValidaForma = "LA LONGITUD DE LA CLAVE NO ES VALIDA."
        
        'registra actividad
        oBitacora.Nombre = NOMBRE_SOLUCION + "-SISTEMA"
        oBitacora.Usuario = NOMBRE_SOLUCION
        oBitacora.registra ("LA LONGITUD DE LA CLAVE NO ES VALIDA.")
        
        Exit Function
    End If
    
    'El primer caracter debe Alfabetico y Mayuscula
    
    'El último debe ser número
'    If Not IsNumeric(Right(strClave, 1)) Then
'        fnValidaForma = "CLAVE NO ES VALIDA."
        
        'registra actividad
'        oBitacora.Nombre = NOMBRE_SOLUCION + "-SISTEMA"
'        oBitacora.Usuario = NOMBRE_SOLUCION
'        oBitacora.registra ("CLAVE NO ES VALIDA.")
        
'        Exit Function
'    End If
    
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbEnter Then
        cmdAceptar_Click
    End If
    
End Sub

Public Function GeneraDigitoVerificador(Referencia As String) As Integer

   Dim x As Integer
   Dim PosImpar As Integer
   Dim PosPar As Integer
   Dim SumaPosiciones As Integer
   Dim TDecenas(1, 9) As Integer
   Dim V2(0, 16) As String
   
   On Error GoTo ErrorHandler
   
   PosPar = 1
   PosImpar = 0
   
   For x = 0 To Len(Referencia) - 1 Step 2
   
      V2(0, PosImpar) = Mid(Referencia, PosImpar + 1, 1) * 1
      
      If (Mid(Referencia, PosPar + 1, 1) * 2) >= 10 Then
         
         Dim RegistroPar As Integer
         RegistroPar = (Mid(Referencia, PosPar + 1, 1) * 2)
         V2(0, PosPar) = CInt(Mid(RegistroPar, 1, 1)) + CInt(Mid(RegistroPar, 2, 1))
      
      Else
          
          V2(0, PosPar) = Mid(Referencia, PosPar + 1, 1) * 2
      
      End If
      
      SumaPosiciones = SumaPosiciones + (CInt(V2(0, PosPar)) + CInt(V2(0, PosImpar)))
      
      PosPar = PosPar + 2
      PosImpar = PosImpar + 2
          
   Next

   If SumaPosiciones > 100 Then
      
      SumaPosiciones = SumaPosiciones - 100
   
   End If
   
   'MsgBox SumaPosiciones
   
   Dim nRow As Integer
   
   For nRow = 0 To 9
      
      If Len(CStr(nRow + 1)) = 2 Then
      
         TDecenas(0, nRow) = Mid((nRow + 1), 2, 1)
      
      Else
         
         TDecenas(0, nRow) = nRow + 1
      
      End If
                
      TDecenas(1, nRow) = (nRow + 1) * 10
      
      If CInt(TDecenas(1, nRow)) >= SumaPosiciones Then
         
         GeneraDigitoVerificador = CInt(TDecenas(1, nRow)) - SumaPosiciones
         
         If GeneraDigitoVerificador = 10 Then
            
            GeneraDigitoVerificador = 0
         
         End If
         
         Exit For
      
      End If
      
   Next

  Exit Function

ErrorHandler:
   
   GeneraDigitoVerificador = -1
   MsgBox Err.Description, vbCritical, "Error No. " & Err.Number & " en Módulo GeneraDigitoVerificador."

End Function
