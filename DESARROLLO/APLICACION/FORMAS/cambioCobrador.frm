VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cambioCobrador 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¿Quien descansa?"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "cambioCobrador.frx":0000
   ScaleHeight     =   1950
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSFrame SSFrame1 
      Height          =   1245
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2196
      _Version        =   196609
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cobrador:"
      Begin VB.ComboBox cbCobradorDescansa 
         BackColor       =   &H8000000E&
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cbCobradorSuple 
         BackColor       =   &H8000000E&
         Height          =   315
         Left            =   2850
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Descansa:"
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Suple:"
         Height          =   255
         Left            =   2850
         TabIndex        =   4
         Top             =   450
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H8000000E&
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1410
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Cobrador igual, No se cambia la ruta."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1530
      Width           =   3465
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Cobrador diferente cambio de ruta"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1260
      Width           =   3465
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Nota:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1260
      Width           =   465
   End
End
Attribute VB_Name = "cambioCobrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strCobradorDescansa As String
Private m_strCobradorSuple As String
Private m_bCancelado As Boolean

Public Property Let strCobradorDescansa(ByVal vData As String)
    m_strCobradorDescansa = vData
End Property

Public Property Get strCobradorDescansa() As String
    strCobradorDescansa = m_strCobradorDescansa
End Property

Public Property Let strCobradorSuple(ByVal vData As String)
    m_strCobradorSuple = vData
End Property

Public Property Get strCobradorSuple() As String
    strCobradorSuple = m_strCobradorSuple
End Property

Public Property Let bCancelado(ByVal vData As Boolean)
    m_bCancelado = vData
End Property

Public Property Get bCancelado() As Boolean
    bCancelado = m_bCancelado
End Property

Private Sub Form_Load()
        
    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        fnLlenaComboCollecion cbCobradorDescansa, oUsuario.cDatos, 0, ""
        fnLlenaComboCollecion cbCobradorSuple, oUsuario.cDatos, 0, ""
    End If
    Set oUsuario = Nothing
    
    m_bCancelado = True
    
End Sub

Private Sub cbCobradorDescansa_Click()
    m_strCobradorDescansa = cbCobradorDescansa.Text
End Sub

Private Sub cbCobradorSuple_Click()
    m_strCobradorSuple = cbCobradorSuple.Text
End Sub

Private Sub cmdAceptar_Click()
    
    If strCobradorDescansa = strCobradorSuple Then
    
        If vbNo = MsgBox("Los cobradores son el mismo, con esta opción NO HAY CAMBIO DE RUTA. ¿Es correcta su elección?", vbQuestion + vbYesNo) Then
            cbCobradorDescansa.SetFocus
            Exit Sub
        End If
          
    End If
        
    Dim strFecha As String
    
    strFecha = Format(Now(), "dd/mm/yyyy")
    'strFecha = DateAdd("d", 1, Format(Now(), "dd/mm/yyyy"))
            
    Dim oPago As New Pago
    Call oPago.asignaRuta(strFecha, m_strCobradorDescansa, m_strCobradorSuple)
    Set oPago = Nothing
    
    m_bCancelado = False
    
    Unload Me
        
End Sub

