VERSION 5.00
Begin VB.Form usuarioPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   585
      Left            =   2610
      TabIndex        =   2
      Top             =   1170
      Width           =   1425
   End
   Begin VB.ComboBox cmbUsuario 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Text            =   "cmbUsuario"
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "¿De que usuario son los pagos?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   4695
   End
End
Attribute VB_Name = "usuarioPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strNombreUsuario As String

'Private Sub cmbUsuario_Click()
'
'    strNombreUsuario = cmbUsuario.Text
'
'End Sub

Private Sub cmdAceptar_Click()
    
    If cmbUsuario.ListIndex = -1 Then
        MsgBox "Seleccione usuario", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    strNombreUsuario = cmbUsuario.Text
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        fnLlenaComboCollecion cmbUsuario, oUsuario.cDatos, 0, ""
        cmbUsuario.ListIndex = 0
    End If
    Set oUsuario = Nothing

End Sub
