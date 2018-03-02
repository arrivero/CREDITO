VERSION 5.00
Begin VB.Form frmsobrantes 
   Caption         =   "Sobrantes"
   ClientHeight    =   3105
   ClientLeft      =   4530
   ClientTop       =   3345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   3870
   Begin VB.TextBox txtfecha 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton optsobrante 
      Caption         =   "Sobrante"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optfaltante 
      Caption         =   "Faltante"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtcantidad 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   390
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1695
      Width           =   735
   End
End
Attribute VB_Name = "frmsobrantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nombreglobal As String
Dim bgral As Integer


Private Sub cmdagregar_Click()
    Dim tipo As String
    Dim cantidad As Long
    If Not IsNumeric(txtcantidad.Text) Then
        MsgBox "El dato de la cantidad no es un número", vbCritical, "Sobrantes"
        txtcantidad.Text = 0
        txtcantidad.SetFocus
    Else
    If optsobrante.Value = True Then
        tipo = "S"
    Else
        If optfaltante.Value = True Then
            tipo = "F"
        End If
    End If
    cantidad = CLng(txtcantidad.Text)
    base.Execute "insert into sobrantes (fecha,tipo,cantidad,usuario) values('" + Format(Now(), "dd/mm/yyyy") + "','" + tipo + "'," + CStr(cantidad) + ",'" + nombreaux + "')"
    MsgBox "El dato fue agregado", vbOKOnly, "Sobrantes"
    End If
    

End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtfecha.Text = Format(Now, "dd/mm/yyyy")
txtcantidad.Text = 0

frmsobrantes.Caption = "Sobrantes " + UCase(nombreaux)

End Sub
