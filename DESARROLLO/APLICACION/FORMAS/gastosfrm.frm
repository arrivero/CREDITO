VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form gastosfrm 
   Caption         =   "Gastos"
   ClientHeight    =   7125
   ClientLeft      =   3120
   ClientTop       =   1035
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   5595
   Begin VB.CommandButton cmdMuestra 
      Caption         =   "Mostrar Gastos"
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gastos Registrados"
      Height          =   4935
      Left            =   30
      TabIndex        =   9
      Top             =   1590
      Width           =   5535
      Begin FPSpread.vaSpread sprGastos 
         Height          =   1575
         Left            =   270
         TabIndex        =   16
         Top             =   2100
         Width           =   5025
         _Version        =   196608
         _ExtentX        =   8864
         _ExtentY        =   2778
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         SpreadDesigner  =   "gastosfrm.frx":0000
      End
      Begin VB.CommandButton cmdgraba 
         Caption         =   "Registra Gastos"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   4200
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid grdgastos 
         Height          =   1605
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2831
         _Version        =   393216
         Cols            =   5
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gastos"
      Height          =   1575
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   5535
      Begin Crystal.CrystalReport reppagos 
         Left            =   3630
         Top             =   1020
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "C:\Facturas\repgastos.rpt"
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtgasto 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "gastosfrm.frx":18DC
         Left            =   4200
         List            =   "gastosfrm.frx":18DE
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox combo 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdagrega 
         Caption         =   "Agrega Gasto"
         Height          =   495
         Left            =   4080
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtfechagasto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbgasto 
         Height          =   315
         ItemData        =   "gastosfrm.frx":18E0
         Left            =   4200
         List            =   "gastosfrm.frx":18E2
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtimporte 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Importe:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   760
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Gasto:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   400
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4230
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "gastosfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_DESCRIPCION = 1
Private Const COL_IMPORTE = 2
Private Const COL_FECHA = 3
Private Const COL_USUARIO = 4

Dim renglongto As Integer

Private Sub Form_Load()
    
'    Me.Caption = "Gastos " + UCase(nombreaux)
'    Dim datos As Recordset
'    renglongto = 0
'    grdgastos.Row = 0
'    grdgastos.Col = 1
'    grdgastos.Text = "Descripción"
'    grdgastos.Col = 2
'    grdgastos.Text = "Importe"
'    grdgastos.Col = 3
'    grdgastos.Text = "Fecha"
'    grdgastos.Col = 4
'    grdgastos.Text = "Usuario"
    
    'Carga el catálgo de gastos
    Dim oGasto As New Gasto
    oGasto.catalogo
    Call llenaComboSpread(sprGastos, COL_DESCRIPCION, oGasto.cDatos, 1)
    Set oGasto = Nothing

'    Set datos = base.OpenRecordset("select * from gasto")
'    datos.MoveFirst
'    While Not datos.EOF
'
'        cmbgasto.AddItem (CStr(datos!gastoid) + " " + datos!descripcion)
'        combo.AddItem (CStr(datos!gastoid))
'        Combo1.AddItem (datos!descripcion)
'        datos.MoveNext
'
'    Wend
'    datos.Close
    
    'cmbgasto.ListIndex = 0
    'combo.ListIndex = 0
    'Combo1.ListIndex = 0
    txtfechagasto.Text = Format(Now, "dd/mm/yyyy")
    
End Sub

Private Sub cmbgasto_Click()
combo.ListIndex = cmbgasto.ListIndex
Combo1.ListIndex = cmbgasto.ListIndex
End Sub

Private Sub cmdagrega_Click()

    Dim i As Integer
    Dim folio, Cliente As Long
    Dim totadeudo, totpago, importe As Double
    Dim gsto As String

    totpago = 0
    totadeudo = 0

    If renglongto = 0 Then
        If IsNumeric(txtimporte.Text) Then
            If txtimporte.Text <> "" And IsNumeric(txtimporte.Text) And CDbl(txtimporte.Text) > 0 And txtgasto.Text <> "" Then
                grdgastos.Row = grdgastos.Rows - 1
                grdgastos.Rows = grdgastos.Row + 2
        '        grdgastos.Col = 1
        '        grdgastos.Text = combo.Text
                grdgastos.Col = 1
                grdgastos.Text = txtgasto.Text
                grdgastos.Col = 2
                grdgastos.Text = Format(txtimporte.Text, "###,###,###,###0.00")
                grdgastos.Col = 3
                grdgastos.Text = txtfechagasto.Text
                grdgastos.Col = 4
                grdgastos.Text = nombreaux
                cmdgraba.Enabled = True
            Else
                MsgBox "El dato del importe es incorrecto", vbCritical, "Registro de Pagos"
                txtimporte.SetFocus
            End If
        End If
    Else
        grdgastos.Row = renglongto
        grdgastos.Col = 1
        If grdgastos.Text <> "" Then
            If IsNumeric(txtimporte.Text) Then
                If txtimporte.Text <> "" And IsNumeric(txtimporte.Text) And CDbl(txtimporte.Text) > 0 And txtgasto.Text <> "" Then
                    grdgastos.Row = renglongto
                    grdgastos.Col = 1
                    grdgastos.Text = txtgasto.Text
                    grdgastos.Col = 2
                    grdgastos.Text = Format(txtimporte.Text, "###,###,###,###0.00")
            '        grdgastos.Col = 3
            '        grdgastos.Text = txtfecha.Text
                    cmdgraba.Enabled = True
                Else
                    If CDbl(txtimporte.Text) = 0 Then
                        For i = renglongto To grdgastos.Rows - 2
                            If i < grdgastos.Rows - 2 Then
                                grdgastos.Row = i + 1
                            Else
                                grdgastos.Row = i
                            End If
                            grdgastos.Col = 1
                            Gasto = grdgastos.Text
                            grdgastos.Col = 2
                            importe = CDbl(Format(grdgastos.Text, "###,###,###,###0.00"))
                    '        grdgastos.Col = 3
                    '        grdgastos.Text = txtfecha.Text
                            grdgastos.Row = i
                            grdgastos.Col = 1
                            grdgastos = Gasto
                            grdgastos.Col = 2
                            grdgastos = Format(importe, "###,###,###,###0.00")
                            grdgastos.Col = 3
                            grdgastos.Text = txtfechagasto.Text
                            cmdgraba.Enabled = True
                        Next i
                        grdgastos.Row = i - 1
                        grdgastos.Col = 1
                        grdgastos = ""
                        grdgastos.Col = 2
                        grdgastos = ""
                        grdgastos.Col = 3
                        grdgastos.Text = ""
                        grdgastos.Rows = grdgastos.Rows - 1
                        If grdgastos.Rows - 2 = 0 Then
                            cmdgraba.Enabled = False
                        End If
                    Else
                        MsgBox "El dato del importe es incorrecto", vbCritical, "Registro de Pagos"
                    End If
            
                txtimporte.SetFocus
                End If
            End If
        End If
    End If
    
    'cmdagregar.Enabled = False
    txtimporte.Text = ""
    txtfechagasto.Text = Format(Now, "dd/mm/yyyy")
    cmbgasto.ListIndex = 0
    combo.ListIndex = 0
    Combo1.ListIndex = 0
    txtgasto.Text = ""
    txtgasto.SetFocus
    renglongto = 0

End Sub

Private Sub cmdgraba_Click()

Dim j, Gasto As Integer
Dim importe As Double
Dim descripcion As String
Dim Usuario As String

    For j = 1 To grdgastos.Rows - 2
        grdgastos.Row = j
'        grdgastos.Col = 1
'        gasto = CDbl(grdgastos.Text)
        grdgastos.Col = 1
        descripcion = grdgastos.Text
        grdgastos.Col = 2
        importe = CDbl(grdgastos.Text)
        grdgastos.Col = 3
        fecha = CDate(grdgastos.Text)
        grdgastos.Col = 4
        Usuario = grdgastos.Text
        'base.Execute "insert into gastos_dia (gasto,fecha,importe,descripcion) values(" + CStr(gasto) + ",'" + Format(fecha, "dd/mm/yyyy") + "','" + CStr(importe) + "','" + descripcion + "')"
        Base.Execute "insert into gastos_dia (gasto,fecha,importe,descripcion,usuario) values(" + CStr(0) + ",'" + Format(fecha, "dd/mm/yyyy") + "','" + CStr(importe) + "','" + descripcion + "','" + Usuario + "')"
    Next j
        
    MsgBox "Los gastos han sido registrados", vbInformation, "Gastos"
    cmdgraba.Enabled = False

End Sub

Private Sub cmdMuestra_Click()
  Base.Execute "delete from rep_gastos_dia"

Dim Gasto As Integer
Dim fecha As Date
Dim importe As Long
Dim descripcion As String
Gasto = 0
For i = 1 To grdgastos.Rows - 2
    grdgastos.Row = i
    grdgastos.Col = 1
    descripcion = grdgastos.Text
    grdgastos.Col = 2
    importe = CLng(grdgastos.Text)
    grdgastos.Col = 3
    fecha = CDate(grdgastos.Text)
    grdgastos.Col = 4
    Usuario = grdgastos.Text
    Base.Execute "insert into rep_gastos_dia (gasto,fecha,importe,descripcion,usuario) values(" + CStr(Gasto) + ",'" + Format(fecha, "dd/mm/yyyy") + "','" + CStr(importe) + "','" + descripcion + "','" + Usuario + "')"
    Gasto = Gasto + 1
Next i
reppagos.PrintReport


End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub


Private Sub grdgastos_DblClick()
renglongto = grdgastos.Row
grdgastos.Col = 1
If grdgastos.Text <> "" Then
    txtgasto.Text = grdgastos.Text
    grdgastos.Col = 2
    txtimporte.Text = Format(CDbl(grdgastos.Text), "###,###,###,###0.00")
    grdgastos.Col = 3
    txtfechagasto.Text = grdgastos.Text
End If

End Sub
