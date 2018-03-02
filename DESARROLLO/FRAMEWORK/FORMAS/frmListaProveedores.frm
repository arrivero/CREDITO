VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form listaProveedoresfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PROVEEDORES"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   465
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2325
   End
   Begin FPSpread.vaSpread sprLista 
      Height          =   7965
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   7095
      _Version        =   196608
      _ExtentX        =   12515
      _ExtentY        =   14049
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmListaProveedores.frx":0000
      Appearance      =   1
      ScrollBarTrack  =   3
   End
End
Attribute VB_Name = "listaProveedoresfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mColUno As String
Public mColDos As String
Public mColTres As String
Public mColCuatro As String
Public mColCinco As String
Public mColSeis As String

Public iProveedor As Integer
Public iSalon As Integer

Public bSeleccion As Boolean

Const COL_UNO = 1
Const COL_DOS = 2
'Const COL_TRES = 3

Private Sub Form_Load()
    
    Dim oProveedor As New cProveedor
    If oProveedor.fnConsultaCatalogo(iSalon) Then
        
        Call fnLlenaTablaCollection(sprLista, oProveedor.cDatos)
        
    End If
    
End Sub

Private Sub sprLista_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    'If OPERACION_CAMBIOS_PROVEEDOR = True Then
        sprLista.Col = 2
        sprLista.Row = Row
        
        If sprLista.Text <> "" Then
            Dim strTemp As String
            frmProveedorMan.Caption = "Cambios al Proveedor"
            frmProveedorMan.tbProveedor.Tag = CAMBIO
            
            sprLista.Col = 1
            sprLista.Row = Row
            
            frmProveedorMan.Tag = Val(sprLista.Text)  'Id del proveedor
            frmProveedorMan.Show vbModal
            strTemp = txtBuscar.Text
            txtBuscar.Text = ""
            txtBuscar.Text = strTemp
            
        End If
    'End If
    
End Sub

Private Sub sprLista_KeyPress(KeyAscii As Integer)
    
    sprLista.Row = sprLista.ActiveRow
    sprLista.Col = COL_UNO
    mColUno = sprLista.Text
    
    sprLista.Col = COL_DOS
    mColDos = sprLista.Text
    
    'sprLista.Col = COL_TRES
    'mColTres = sprLista.Text
    
    'sprLista.Col = COL_CUATRO
    'mColCuatro = sprLista.Text
    
    'sprLista.Col = COL_CINCO
    'mColCinco = sprLista.Text
    
    'sprLista.Col = COL_SEIS
    'mColSeis = sprLista.Text
        
    bSeleccion = True

    Unload Me
    
End Sub

'Private Sub txtBuscar_Change()
'
'    If txtBuscar.Text <> "" Then
'
'        Dim oCliente As New cCliente
'        Dim cPartidas As New Collection
'        If oCliente.clienteBusca(txtBuscar.Text) Then
'            Set cPartidas = oCliente.cDatos
'
'            sprLista.Row = 1
'            sprLista.Col = 1
'            sprLista.Row2 = -1
'            sprLista.Col2 = -1
'
'            sprLista.BlockMode = True
'            sprLista.Action = ActionClearText
'            sprLista.BlockMode = False
'
'            Call fnLlenaTablaCollection(sprLista, cPartidas)
'        End If
'
'    End If
'End Sub
