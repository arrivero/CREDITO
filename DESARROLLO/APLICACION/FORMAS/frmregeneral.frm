VERSION 5.00
Begin VB.Form reGeneralfrm 
   Caption         =   "Resumen Diario General"
   ClientHeight    =   4065
   ClientLeft      =   4530
   ClientTop       =   3345
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   3870
   Begin VB.TextBox txtdevo 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "0"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtcheque 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "0"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtefeche 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton optsobrante 
      Caption         =   "Sobrante"
      Enabled         =   0   'False
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optfaltante 
      Caption         =   "Faltante"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtfecha 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtcantidad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdreporte 
      Caption         =   "Generar Reporte"
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Devoluciones:"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   2895
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Cheque:"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   2535
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Efectivo:"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2175
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3720
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   270
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   1455
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3720
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "reGeneralfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nombreglobal As String
Dim bgral As Integer


Private Sub repgeneral()

Dim xlaCPP As Object 'Aplicacion
Dim xlwCPP As Object 'archivo
Dim xlsCPP As Object 'Worksheet
        
Dim bdd As DataBase
Dim r As Recordset
Dim a, b, FechaInicio, archivo As String
Dim i, j, k, Fila, Columna As Integer
Dim res As Integer
Dim cdtotal As Double

archivo = "c:\facturas\Reporte\General"
''archivo = archivo + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")" + ".xls"
'Set bdd = CurrentDb
'Set xlwCPP = CreateObject("Excel.Sheet.8")
'Set xlwCPP = GetObject(archivo)
'xlwCPP.Application.Visible = True
'xlwCPP.Parent.Windows(1).Visible = True
       
'Crear el archivo de Excel
Set xlwCPP = CreateObject("Excel.Sheet.8")
Set xlsCPP = xlwCPP.Activesheet
Set xlaCPP = xlsCPP.Parent.Parent
'xlwCPP.SaveAs (archivo)
xlwCPP.Application.Visible = True
xlaCPP.ActiveWindow.WindowState = -4137 'Maximiza la ventana
    
'xlsCPP.Name = "General"
xlsCPP.Name = "General" + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")"

'''''''''''''''''''''''''''''''''''''''''' Resumen
Fila = 5
Columna = 1
j = 5
Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 5")
Set xlsCPP = xlwCPP.Worksheets("General" + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")")
xlwCPP.Worksheets("General" + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")").Activate
Set xlsCPP = xlwCPP.Activesheet


xlaCPP.Cells(1, 1).Value = "Resumen Diario"
xlaCPP.Range("A1:B1").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 12
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(1, 4).Value = "Fecha: "
xlaCPP.Range("D1").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = False
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 12
xlaCPP.Selection.Font.Bold = True


xlaCPP.Cells(2, 4).Value = "Total Cobrado:"
xlaCPP.Range("D2:E2").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

xlaCPP.Cells(3, 4).Value = "'-Facturas:"
xlaCPP.Range("D3:E3").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

xlaCPP.Cells(4, 4).Value = "'-Gastos:"
xlaCPP.Range("D4:E4").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

'xlaCPP.Cells(5, 4).Value = "Total:"
xlaCPP.Cells(5, 4).Value = "+Electrónicos:"
xlaCPP.Range("D5:E5").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

xlaCPP.Cells(6, 4).Value = "'-Faltante:"
xlaCPP.Range("D6:E6").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

xlaCPP.Cells(7, 4).Value = "'+Sobrante:"
xlaCPP.Range("D7:E7").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

xlaCPP.Cells(8, 4).Value = "+Efectivo:"
xlaCPP.Range("D8:E8").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False
'*************************total*********************************************
xlaCPP.Cells(9, 4).Value = "+Cheque:"
xlaCPP.Range("D9:E9").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False
'***********************************************************************
xlaCPP.Cells(10, 4).Value = "+Devoluciones:"
xlaCPP.Range("D10:E10").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

xlaCPP.Cells(11, 4).Value = "Total:"
xlaCPP.Range("D11:E11").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = False

'xlaCPP.Range("F4").Select
'    With xlaCPP.Selection.Borders(4)
'        .LineStyle = xlContinuous
'        .Weight = 3
'        .ColorIndex = xlAutomatic
'    End With
    
xlaCPP.Range("F10").Select
    With xlaCPP.Selection.Borders(4)
        .LineStyle = xlContinuous
        .Weight = 3
        .ColorIndex = xlAutomatic
    End With

xlaCPP.Range("A11:I11").Select
    With xlaCPP.Selection.Borders(4)
        .LineStyle = xlContinuous
        .Weight = 3
        .ColorIndex = xlAutomatic
    End With

While Not r.EOF
    xlaCPP.Cells(1, 5).Value = r!fechard
    
    xlaCPP.Range("E1:F1").Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.MergeCells = True
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 11
    xlaCPP.Selection.Font.Bold = True
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
    
    xlaCPP.Cells(2, 6).Value = Format(r!cobrado, "###,###,###,###0.00")
    xlaCPP.Cells(3, 6).Value = Format(r!facturas, "###,###,###,###0.00")
    xlaCPP.Cells(4, 6).Value = Format(r!gastos, "###,###,###,###0.00")
    'xlaCPP.Cells(5, 6).Value = Format((r!cobrado - r!facturas) - r!gastos, "###,###,###,###0.00")
    xlaCPP.Cells(5, 6).Value = Format(r!electricos, "###,###,###,###0.00")
    xlaCPP.Cells(6, 6).Value = Format(r!faltante, "###,###,###,###0.00")
    xlaCPP.Cells(7, 6).Value = Format(r!sobrante, "###,###,###,###0.00")
    xlaCPP.Cells(8, 6).Value = Format(r!ec, "###,###,###,###0.00")
    xlaCPP.Cells(9, 6).Value = Format(r!cheque, "###,###,###,###0.00")
    xlaCPP.Cells(10, 6).Value = Format(r!devolucion, "###,###,###,###0.00")
    xlaCPP.Cells(11, 6).Value = Format((((r!cobrado - r!facturas) - r!gastos) - r!faltante) + r!sobrante + r!ec + r!cheque + r!devolucion + r!electricos, "###,###,###,###0.00")
    j = j + 1
    r.MoveNext
Wend
r.Close

''''''''''''''''''''''''''''''''''''''''''CORTE DIARIO

xlaCPP.Cells(13, 1).Value = "Corte Diario"
xlaCPP.Range("A13:B13").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 12
xlaCPP.Selection.Font.Bold = True

a = "B"
b = "C"
j = 2
For i = 1 To 3
    xlaCPP.Cells(15, j).Value = "Folio"
    xlaCPP.Range(a + "15").Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.MergeCells = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = True
    
    xlaCPP.Cells(15, j + 1).Value = "Cantidad"
    xlaCPP.Range(b + "15").Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.MergeCells = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = True
    j = j + 3
    If i = 1 Then
        a = "E"
        b = "F"
    Else
        a = "H"
        b = "I"
    End If
Next i

xlaCPP.Cells(13, 4).Value = "Fecha:"
xlaCPP.Range("D13").Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = False
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = True

j = 16
cdtotal = 0
i = 1
Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 1")
If Not r.EOF Then
    xlaCPP.Cells(13, 5).Value = IIf(IsNull(r!fechacd), " ", r!fechacd)
    xlaCPP.Range("E13").Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = True
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
End If

While Not r.EOF
    
    xlaCPP.Cells(j, 2).Value = r!folio1
    xlaCPP.Cells(j, 3).Value = Format(r!pago1, "###,###,###,###0.00")
    xlaCPP.Cells(j, 5).Value = r!folio2
    xlaCPP.Cells(j, 6).Value = Format(r!pago2, "###,###,###,###0.00")
    xlaCPP.Cells(j, 8).Value = r!folio3
    xlaCPP.Cells(j, 9).Value = Format(r!pago3, "###,###,###,###0.00")
    
    j = j + 1
    i = i + 1
    cdtotal = cdtotal + r!pago1 + r!pago2 + r!pago3
    r.MoveNext
    If i = 20 And Not r.EOF Then
        a = "B"
        b = "C"
        k = 2
        For i = 1 To 3
            xlaCPP.Cells(j, k).Value = "Folio"
            xlaCPP.Range(a + CStr(j)).Select
            xlaCPP.Selection.WrapText = False
            xlaCPP.Selection.Orientation = 0
            xlaCPP.Selection.AddIndent = False
            xlaCPP.Selection.MergeCells = False
            xlaCPP.Selection.Font.Name = "Arial"
            xlaCPP.Selection.Font.Size = 10
            xlaCPP.Selection.Font.Bold = True
            
            xlaCPP.Cells(j, k + 1).Value = "Cantidad"
            xlaCPP.Range(b + CStr(j)).Select
            xlaCPP.Selection.WrapText = False
            xlaCPP.Selection.Orientation = 0
            xlaCPP.Selection.AddIndent = False
            xlaCPP.Selection.MergeCells = False
            xlaCPP.Selection.Font.Name = "Arial"
            xlaCPP.Selection.Font.Size = 10
            xlaCPP.Selection.Font.Bold = True
            k = k + 3
            If i = 1 Then
                a = "E"
                b = "F"
            Else
                a = "H"
                b = "I"
            End If
        Next i
        j = j + 1
        i = 1
    End If
    
Wend
r.Close

a = "B"
b = "C"
For i = 1 To 3
    xlaCPP.Range(a + "15:" + b + CStr(j - 1)).Select
        With xlaCPP.Selection.Borders(1)
            .LineStyle = xlContinuous
            .Weight = 2
            .ColorIndex = xlAutomatic
        End With
        With xlaCPP.Selection.Borders(2)
            .LineStyle = xlContinuous
            .Weight = 2
            .ColorIndex = xlAutomatic
        End With
        With xlaCPP.Selection.Borders(4)
            .LineStyle = xlContinuous
            .Weight = 2
            .ColorIndex = xlAutomatic
        End With
        With xlaCPP.Selection.Borders(3)
            .LineStyle = xlContinuous
            .Weight = 2
            .ColorIndex = xlAutomatic
        End With
    If i = 1 Then
        a = "E"
        b = "F"
    Else
        a = "H"
        b = "I"
    End If
Next i

xlaCPP.Cells(j + 1, 6).Value = "Total Corte Diario"
xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
'xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
xlaCPP.Range("H" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

j = j + 2

xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
    With xlaCPP.Selection.Borders(4)
        .LineStyle = xlContinuous
        .Weight = 3
        .ColorIndex = xlAutomatic
    End With

j = j + 2

''''''''''''''''''''''''''''''''''''''''''CRÉDITOS NUEVOS

xlaCPP.Cells(j, 1).Value = "Créditos Nuevos"
xlaCPP.Range("A" + CStr(j) + ":B" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 12
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 4).Value = "Fecha Alta Crédito:"
xlaCPP.Range("D" + CStr(j) + ":E" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = True

Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 2")
If Not r.EOF Then
    xlaCPP.Cells(j, 6).Value = IIf(IsNull(r!Fechac), " ", r!Fechac)
    xlaCPP.Range("F" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = True
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
End If
j = j + 2

xlaCPP.Cells(j, 1).Value = "       Folio"
xlaCPP.Range("A" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 3).Value = "Nombre"
xlaCPP.Range("C" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 6).Value = "      Crédito"
xlaCPP.Range("F" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 8).Value = "     Inicio"
xlaCPP.Range("H" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True


xlaCPP.Cells(j, 9).Value = "       Fin"
xlaCPP.Range("I" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

j = j + 1
cdtotal = 0
While Not r.EOF

    xlaCPP.Cells(j, 1).Value = CStr(r!factura)
    xlaCPP.Range("A" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 3).Value = r!Nombre
    xlaCPP.Range("C" + CStr(j) + ":D" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 5).Value = IIf(r!elec = 1, "Electricos", "")
    xlaCPP.Range("E" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 6).Value = Format(r!credito, "###,###,###,###0.00")
    xlaCPP.Range("F" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 8).Value = r!fechaini
    xlaCPP.Range("H" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"

    xlaCPP.Cells(j, 9).Value = r!fechatermina
    xlaCPP.Range("I" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"

    j = j + 1
    i = i + 1
    cdtotal = cdtotal + r!credito
    r.MoveNext
    If i = 20 And Not r.EOF Then
        a = "B"
        b = "C"
        k = 2
        xlaCPP.Cells(j, 1).Value = "       Folio"
        xlaCPP.Range("A" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 3).Value = "Nombre"
        xlaCPP.Range("C" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 6).Value = "      Crédito"
        xlaCPP.Range("F" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 8).Value = "     Inicio"
        xlaCPP.Range("H" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = False
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 9).Value = "       Fin"
        xlaCPP.Range("I" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True

        j = j + 1
        i = 1
    End If

Wend
r.Close

xlaCPP.Cells(j + 1, 6).Value = "Total Créditos Nuevos"
xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
'xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
xlaCPP.Range("H" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

j = j + 2

xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
    With xlaCPP.Selection.Borders(4)
        .LineStyle = xlContinuous
        .Weight = 3
        .ColorIndex = xlAutomatic
    End With

j = j + 2

''''''''''''''''''''''''''''''''''''''''''CRÉDITOS TERMINADOS

xlaCPP.Cells(j, 1).Value = "Créditos Terminados"
'xlaCPP.Range("A" + CStr(j) + ":B" + CStr(j)).Select
xlaCPP.Range("A" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 12
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 4).Value = "Fecha Terminación Crédito:"
'xlaCPP.Range("D" + CStr(j) + ":E" + CStr(j)).Select
xlaCPP.Range("D" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = True

Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 3")
If Not r.EOF Then
    xlaCPP.Cells(j, 7).Value = IIf(IsNull(r!fechatermina), " ", r!fechatermina)
    xlaCPP.Range("G" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = True
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
End If

j = j + 2

xlaCPP.Cells(j, 1).Value = "       Folio"
xlaCPP.Range("A" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 3).Value = "Nombre"
xlaCPP.Range("C" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 6).Value = "      Crédito"
xlaCPP.Range("F" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 8).Value = "      Alta"
xlaCPP.Range("H" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True


xlaCPP.Cells(j, 9).Value = "     Inicio"
xlaCPP.Range("I" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

j = j + 1
cdtotal = 0
While Not r.EOF

    xlaCPP.Cells(j, 1).Value = CStr(r!factura)
    xlaCPP.Range("A" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 3).Value = r!Nombre
    xlaCPP.Range("C" + CStr(j) + ":E" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 6).Value = Format(r!credito, "###,###,###,###0.00")
    xlaCPP.Range("F" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 8).Value = r!Fechac
    xlaCPP.Range("H" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"

    xlaCPP.Cells(j, 9).Value = r!fechaini
    xlaCPP.Range("I" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"

    j = j + 1
    i = i + 1
    cdtotal = cdtotal + r!credito
    r.MoveNext
    If i = 20 And Not r.EOF Then
        a = "B"
        b = "C"
        k = 2
        xlaCPP.Cells(j, 1).Value = "       Folio"
        xlaCPP.Range("A" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 3).Value = "Nombre"
        xlaCPP.Range("C" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 6).Value = "      Crédito"
        xlaCPP.Range("F" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 8).Value = "      Alta"
        xlaCPP.Range("H" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = False
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 9).Value = "     Inicio"
        xlaCPP.Range("I" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True

        j = j + 1
        i = 1
    End If

Wend
r.Close

xlaCPP.Cells(j + 1, 6).Value = "Total  Terminados"
'xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
xlaCPP.Range("F" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
'xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
xlaCPP.Range("H" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

j = j + 2

xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
    With xlaCPP.Selection.Borders(4)
        .LineStyle = xlContinuous
        .Weight = 3
        .ColorIndex = xlAutomatic
    End With

j = j + 2

''''''''''''''''''''''''''''''''''''''''''GASTOS

xlaCPP.Cells(j, 1).Value = "Gastos"
'xlaCPP.Range("A" + CStr(j) + ":B" + CStr(j)).Select
xlaCPP.Range("A" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 12
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 4).Value = "Fecha:"
'xlaCPP.Range("D" + CStr(j) + ":E" + CStr(j)).Select
xlaCPP.Range("D" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 11
xlaCPP.Selection.Font.Bold = True

Set r = Base.OpenRecordset("SELECT * FROM regen where cve = 4")
If Not r.EOF Then
    xlaCPP.Cells(j, 5).Value = IIf(IsNull(r!Fechag), " ", r!Fechag)
    xlaCPP.Range("E" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = True
    xlaCPP.Selection.numberFormat = "dd/mm/yyyy"
End If

j = j + 2

xlaCPP.Cells(j, 2).Value = "Descripción"
xlaCPP.Range("B" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j, 6).Value = "Importe"
xlaCPP.Range("F" + CStr(j)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

j = j + 1
cdtotal = 0
While Not r.EOF

    xlaCPP.Cells(j, 2).Value = r!descripcion
    xlaCPP.Range("B" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    xlaCPP.Cells(j, 6).Value = Format(r!importe, "###,###,###,###0.00")
    xlaCPP.Range("F" + CStr(j)).Select
    xlaCPP.Selection.WrapText = False
    xlaCPP.Selection.Orientation = 0
    xlaCPP.Selection.AddIndent = False
    xlaCPP.Selection.Font.Name = "Arial"
    xlaCPP.Selection.Font.Size = 10
    xlaCPP.Selection.Font.Bold = False
    xlaCPP.Selection.MergeCells = True
    
    j = j + 1
    i = i + 1
    cdtotal = cdtotal + r!importe
    r.MoveNext
    If i = 20 And Not r.EOF Then
        a = "B"
        b = "C"
        k = 2
        xlaCPP.Cells(j, 2).Value = "Descripción"
        xlaCPP.Range("B" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        xlaCPP.Cells(j, 6).Value = "Importe"
        xlaCPP.Range("F" + CStr(j)).Select
        xlaCPP.Selection.WrapText = False
        xlaCPP.Selection.Orientation = 0
        xlaCPP.Selection.AddIndent = False
        xlaCPP.Selection.MergeCells = True
        xlaCPP.Selection.Font.Name = "Arial"
        xlaCPP.Selection.Font.Size = 10
        xlaCPP.Selection.Font.Bold = True
        
        j = j + 1
        i = 1
    End If

Wend
r.Close

xlaCPP.Cells(j + 1, 6).Value = "Total  Gastos"
'xlaCPP.Range("F" + CStr(j + 1) + ":G" + CStr(j + 1)).Select
xlaCPP.Range("F" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

xlaCPP.Cells(j + 1, 8).Value = Format(cdtotal, "###,###,###,###0.00")
'xlaCPP.Range("H" + CStr(j + 1) + ":I" + CStr(j + 1)).Select
xlaCPP.Range("H" + CStr(j + 1)).Select
xlaCPP.Selection.WrapText = False
xlaCPP.Selection.Orientation = 0
xlaCPP.Selection.AddIndent = False
xlaCPP.Selection.MergeCells = True
xlaCPP.Selection.Font.Name = "Arial"
xlaCPP.Selection.Font.Size = 10
xlaCPP.Selection.Font.Bold = True

j = j + 2

xlaCPP.Range("A" + CStr(j) + ":I" + CStr(j)).Select
    With xlaCPP.Selection.Borders(4)
        .LineStyle = xlContinuous
        .Weight = 3
        .ColorIndex = xlAutomatic
    End With

j = j + 2
bgral = 0
archivo = "C:\General\General"
archivo = archivo + "(" + Format(txtfecha.Text, "dd-mm-yyyy") + ")" + ".xls"
nombreglobal = archivo
On Error GoTo fin1
xlwCPP.Saveas (archivo)
bgral = 1
GoTo fin2
fin1: MsgBox "El archivo generado no va a ser guardado", vbCritical, "Reporte General"

fin2: End Sub
Private Sub resumen()

Dim datos, datos1, datos2 As Recordset
Dim pagos1, gastos, folio, folioe, total_anterior As Double
Dim conta As Integer

If txtfecha.Text <> "" Then
        
    total_anterior = 0
    conta = 0
    Set datos2 = Base.OpenRecordset("select * from total_dia_anterior")
    While Not datos2.EOF
        'If DatePart("d", datos2!fecha) = DatePart("d", CDate(txtfecha.Text)) And DatePart("m", datos2!fecha) = DatePart("m", CDate(txtfecha.Text)) And DatePart("yyyy", datos2!fecha) = DatePart("yyyy", CDate(txtfecha.Text)) Then
        If datos2!fecha = CDate(txtfecha.Text) Then
            datos2.MovePrevious
            conta = 1
        End If
        
        total_anterior = IIf(IsNull(datos2!total_anterior), 0, datos2!total_anterior)
        If conta = 1 Then
            datos2.MoveLast
            datos2.MoveNext
        Else
            datos2.MoveNext
        End If
    Wend
    datos2.Close
        
    
    '*********************************************************************
    If conta = 0 Then
        Set datos2 = Base.OpenRecordset("select * from total_dia_anterior where clng(fecha) < " + CStr(CLng(CDate(txtfecha.Text))) + " order by fecha desc")
        While Not datos2.EOF
            total_anterior = IIf(IsNull(datos2!total_anterior), 0, datos2!total_anterior)
            datos2.MoveLast
            datos2.MoveNext
        Wend
        datos2.Close
    End If
    '*********************************************************************
    
    
    Base.Execute "delete from resumen_diario_gral where cstr(datepart('d',fecha))=" + CStr(DatePart("d", CDate(txtfecha.Text))) + " and cstr(datepart('m',fecha))=" + CStr(DatePart("m", CDate(txtfecha.Text))) + " and cstr(datepart('yyyy',fecha))=" + CStr(DatePart("yyyy", CDate(txtfecha.Text)))
    'base.Execute "delete from resumen_diario"
        
    conta = 0
    Set datos = Base.OpenRecordset("select * from qrysumapagos")
    While Not datos.EOF
        If datos!fecha = CDate(txtfecha.Text) Then
            pagos1 = datos!pagos
            conta = 1
        End If
        datos.MoveNext
    Wend
    datos.Close
    If conta = 0 Then
        pagos1 = 0
    End If

    conta = 0
    Set datos1 = Base.OpenRecordset("select * from qrysumacredito")
    While Not datos1.EOF
        If datos1!fecha = CDate(txtfecha.Text) Then
            folio = datos1!credito
            conta = 1
        End If
        datos1.MoveNext
    Wend
    datos1.Close
    If conta = 0 Then
        folio = 0
    End If

    conta = 0
    Set datos1 = Base.OpenRecordset("select * from qrysumacreditoelectricos")
    While Not datos1.EOF
        If datos1!fecha = CDate(txtfecha.Text) Then
            folioe = datos1!creditoe
            conta = 1
        End If
        datos1.MoveNext
    Wend
    datos1.Close
    If conta = 0 Then
        folioe = 0
    End If
    
    conta = 0
    Set datos2 = Base.OpenRecordset("select * from qrysumagastos")
    While Not datos2.EOF
        If datos2!fecha = CDate(txtfecha.Text) Then
            gastos = datos2!importe
            conta = 1
        End If
        datos2.MoveNext
    Wend
    datos2.Close
    If conta = 0 Then
        gastos = 0
    End If

    
    If txtcantidad.Text = "" Then
        txtcantidad.Text = 0
        If Not IsNumeric(txtcantidad.Text) Then
            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
            txtcantidad.Text = 0
            txtcantidad.SetFocus
        End If
    End If
    If Me.txtefeche.Text = "" Then
        txtefeche.Text = 0
        If Not IsNumeric(txtefeche.Text) Then
            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
            txtefeche.Text = 0
            txtefeche.SetFocus
        End If
    End If
    
    If Me.txtcheque.Text = "" Then
        txtcheque.Text = 0
        If Not IsNumeric(txtcheque.Text) Then
            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
            txtcheque.Text = 0
            txtcheque.SetFocus
        End If
    End If
    
    If Me.txtdevo.Text = "" Then
        txtdevo.Text = 0
        If Not IsNumeric(txtdevo.Text) Then
            'MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario"
            txtdevo.Text = 0
            txtdevo.SetFocus
        End If
    End If
    
    If optsobrante.Value = True Then
        Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,fechard,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,electricos) values(" + CStr(5) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(0) + "," + CStr(txtcantidad.Text) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(folioe) + ")"
        'base.Execute "insert into regen (cve,fechard,cobrado,facturas,gastos,faltante,sobrante) values(" + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(0) + "," + CStr(txtcantidad.Text) + ")"
        Base.Execute "insert into resumen_diario_gral (fecha,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,anterior) values('" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(0) + "," + CStr(txtcantidad.Text) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(total_anterior) + ")"
    Else
        Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,fechard,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,electricos) values(" + CStr(5) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(txtcantidad.Text) + "," + CStr(0) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(folioe) + ")"
        'base.Execute "insert into regen (cve,fechard,cobrado,facturas,gastos,faltante,sobrante) values(" + CStr(5) + ",'" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(txtcantidad.Text) + "," + CStr(0) + ")"
        Base.Execute "insert into resumen_diario_gral (fecha,cobrado,facturas,gastos,faltante,sobrante,ec,cheque,devolucion,anterior) values('" + CStr(txtfecha.Text) + "'," + CStr(pagos1) + "," + CStr(folio) + "," + CStr(gastos) + "," + CStr(txtcantidad.Text) + "," + CStr(0) + "," + CStr(txtefeche.Text) + "," + CStr(txtcheque.Text) + "," + CStr(txtdevo.Text) + "," + CStr(total_anterior) + ")"
    End If
    'rediario.PrintReport
    
    
    total_anterior = 0
    conta = 0
    'Set datos2 = base.OpenRecordset("select * from resumen_por_fecha where DatePart('d', fecha) >= " + CStr(DatePart("d", CDate(txtfecha.Text))) + " And DatePart('m', fecha) >= " + CStr(DatePart("m", CDate(txtfecha.Text))) + " And DatePart('yyyy', fecha) >= " + CStr(DatePart("yyyy", CDate(txtfecha.Text))) + " order by fecha asc")
    Set datos2 = Base.OpenRecordset("select * from resumen_por_fecha where clng(fecha) >= " + CStr(CLng(CDate(txtfecha.Text))) + " order by fecha asc")
    While Not datos2.EOF
        'If DatePart("d", datos2!fecha) = DatePart("d", CDate(txtfecha.Text)) And DatePart("m", datos2!fecha) = DatePart("m", CDate(txtfecha.Text)) And DatePart("yyyy", datos2!fecha) = DatePart("yyyy", CDate(txtfecha.Text)) Then
        If conta = 1 Then
            Base.Execute "update resumen_diario_gral set anterior=" + CStr(total_anterior) + " where datepart('d',fecha)=" + CStr(DatePart("d", datos2!fecha)) + " and datepart('m',fecha)=" + CStr(DatePart("m", datos2!fecha)) + " and datepart('yyyy',fecha)=" + CStr(DatePart("yyyy", datos2!fecha))
        End If
        If datos2!fecha = CDate(txtfecha.Text) Then
            conta = 1
            total_anterior = (((((((datos2!anterior_ + datos2!cobrado) - datos2!facturas) - datos2!gastos) - datos2!faltante) + datos2!sobrante) + datos2!cheques) - datos2!depositos) - datos2!gtos + datos2!devo + datos2!electricos
        Else
            total_anterior = (((((((total_anterior + datos2!cobrado) - datos2!facturas) - datos2!gastos) - datos2!faltante) + datos2!sobrante) + datos2!cheques) - datos2!depositos) - datos2!gtos + datos2!devo + datos2!electricos
        End If
        datos2.MoveNext
    Wend
    datos2.Close
   
Else
'    MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario"
'    txtfecha.Text = Format(Now, "dd/mm/yyyy")
'    txtfecha.SetFocus
End If



End Sub
Private Sub cd()

Dim i, j As Integer
Dim datos As Recordset
Dim folio1, folio2, folio3 As Long
Dim pago1, pago2, pago3 As Double
Dim fecha As Date

folio1 = 0
folio2 = 0
folio3 = 0

pago1 = 0
pago2 = 0
pago3 = 0

j = O

'base.Execute "delete from cortediario"

Set datos = Base.OpenRecordset("select * from cdiario")
'datos.MoveFirst
While Not datos.EOF
    For i = 1 To 3
        If datos.EOF Then
            GoTo fincd
        Else
            Select Case i
            Case 1
                folio1 = datos!factura
                pago1 = datos!Cantpagada
            Case 2
                folio2 = datos!factura
                pago2 = datos!Cantpagada
            Case 3
                folio3 = datos!factura
                pago3 = datos!Cantpagada
            End Select
        End If
        fecha = datos!fecha
        datos.MoveNext
    Next i
fincd:
    Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,fechacd,folio1,pago1,folio2,pago2,folio3,pago3) values(" + CStr(1) + "," + CStr(1) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + ",'" + CStr(fecha) + "'," + CStr(folio1) + "," + CStr(pago1) + "," + CStr(folio2) + "," + CStr(pago2) + "," + CStr(folio3) + "," + CStr(pago3) + ")"
    'base.Execute "insert into regen (cve,fechacd,folio1,pago1,folio2,pago2,folio3,pago3) values(" + CStr(1) + ",'" + CStr(fecha) + "'," + CStr(folio1) + "," + CStr(pago1) + "," + CStr(folio2) + "," + CStr(pago2) + "," + CStr(folio3) + "," + CStr(pago3) + ")"
    j = j + 1
    folio1 = 0
    folio2 = 0
    folio3 = 0
    pago1 = 0
    pago2 = 0
    pago3 = 0
Wend

datos.Close
'If j > 0 Then
'    diario.PrintReport
'Else
'    MsgBox "No existen pagos registrados para el dia de hoy", vbInformation, "Corte Diario"
'End If
End Sub
Private Sub nuevos()
Dim datos, datos1, datos2 As Recordset
Dim pagos1, gastos, folio As Double
Dim conta As Integer

If txtfecha.Text <> "" Then
        
    'base.Execute "delete from nuevos"
        
    conta = 0
    Set datos = Base.OpenRecordset("select * from qrycreditos")
    While Not datos.EOF
        If datos!fecha = CDate(txtfecha.Text) Then
            Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fechac,Status,elec) values(" + CStr(2) + "," + CStr(0) + "," + CStr(2) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!credito) + ",'" + CStr(datos!Nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "'," + CStr(IIf(IsNull(datos!electrico), 0, datos!electrico)) + ")"
            'base.Execute "insert into regen (cve,No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fechac,Status) values(" + CStr(2) + "," + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!Credito) + ",'" + CStr(datos!nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "')"
            conta = 1
        End If
        datos.MoveNext
    Wend
    datos.Close
'    If conta = 1 Then
'        renuevo.PrintReport
'
'    Else
'        MsgBox "No existen créditos nuevos registrados para el dia seleccionado", vbInformation, "Resumen Diario"
'        txtfecha.SetFocus
'    End If
   'reporte_termina
Else
'    MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario"
'    txtfecha.Text = Format(Now, "dd/mm/yyyy")
'    txtfecha.SetFocus
End If
End Sub
Private Sub termina()
Dim datos, datos1, datos2 As Recordset
Dim pagos1, gastos, folio As Double
Dim conta As Integer

If txtfecha.Text <> "" Then
        
    'base.Execute "delete from terminados"
        
    conta = 0
    Set datos = Base.OpenRecordset("select * from qrycreditosterminados ")
    While Not datos.EOF
        If datos!fechatermina = CDate(txtfecha.Text) Then
            Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fechac,Status) values(" + CStr(3) + "," + CStr(0) + "," + CStr(0) + "," + CStr(3) + "," + CStr(0) + "," + CStr(0) + "," + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!credito) + ",'" + CStr(datos!Nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "')"
            'base.Execute "insert into regen (cve,No_cliente,Factura,Credito,Nombre,Cantpagar,Financiamiento,Canttotal,No_pagos,Fechaini,Fechatermina,fechac,Status) values(" + CStr(3) + "," + CStr(datos!no_cliente) + "," + CStr(datos!factura) + "," + CStr(datos!Credito) + ",'" + CStr(datos!nombre) + "'," + CStr(datos!Cantpagar) + "," + CStr(datos!financiamiento) + "," + CStr(datos!Canttotal) + "," + CStr(datos!no_pagos) + ",'" + Format(datos!fechaini, "dd/mm/yyyy") + "','" + Format(datos!fechatermina, "dd/mm/yyyy") + "','" + Format(datos!fecha, "dd/mm/yyyy") + "','" + datos!Status + "')"
            conta = 1
        End If
        datos.MoveNext
    Wend
    datos.Close
'    If conta = 1 Then
'        retermina.PrintReport
'    Else
'        MsgBox "No existen créditos terminados para el dia seleccionado", vbInformation, "Resumen Diario"
'        txtfecha.SetFocus
'    End If

Else
'    MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario"
'    txtfecha.Text = Format(Now, "dd/mm/yyyy")
'    txtfecha.SetFocus
End If
End Sub
Private Sub cmdreporte_Click()
Dim datos As Recordset
Dim datos2 As Recordset
Dim datos3 As Recordset
Dim fecha As Date
Dim usuario As String
Dim Aux As Long
Dim i As Long
i = 0
If txtfecha.Text <> "" Then
    If Not IsNumeric(txtcantidad.Text) Then
        MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario General"
        txtcantidad.Text = 0
        txtcantidad.SetFocus
    Else
        Set datos = Base.OpenRecordset("cuidado3")
        ordensigue = 0
        While Not datos.EOF
        usuario = LCase(CStr(datos!estado))
        If usuario = "nuevo leon" Or usuario = "nuevo león" Or usuario = "nl" Or usuario = "nulo" Then
        usuario = "general"
        End If
            Base.Execute "insert into pagos (no_cliente,factura,fecha,Cantpagada,Cantadeudada,orden,cons_pago,usuario,hora,lugar) values(" + CStr(datos!no_cliente) + "," + CStr(datos!factura) + ",'" + Format(Now(), "dd/mm/yyyy") + "'," + CStr(0) + "," + CStr(datos!adeudo) + "," + CStr(ordensigue) + "," + CStr(1) + ",'" + usuario + "','33:33 AM','Captura')"
            ordensigue = ordensigue + 1
            datos.MoveNext
        Wend
        datos.Close
        
        
        Base.Execute "delete from regen"
        Call cd
        Call nuevos
        Call termina
        Call gastos
        Call resumen
        Call repgeneral
        If bgral = 1 Then
            MsgBox "El reporte general se encuenta en " + nombreglobal, vbInformation, "Resumen Diario General"
        End If
    End If
Else
    MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario General"
    txtfecha.Text = Format(Now, "dd/mm/yyyy")
    txtfecha.SetFocus
End If



'Dim datos As Recordset
'Dim fecha As Date
'
'If txtfecha.Text <> "" Then
'    If Not IsNumeric(txtcantidad.Text) Then
'        MsgBox "El dato de la cantidad es incorrecto", vbCritical, "Resumen Diario General"
'        txtcantidad.Text = 0
'        txtcantidad.SetFocus
'    Else
'
'        Set datos = base.OpenRecordset("select * from adeudo_pv where cdbl(fecha)<=" + CStr(CDbl(CDate(Format(Now(), "dd/mm/yyyy")))))
'        While Not datos.EOF
'            base.Execute "insert into pagos (no_cliente,factura,fecha,Cantpagada,Cantadeudada,orden) values(" + CStr(datos!no_cliente) + "," + CStr(datos!factura) + ",'" + Format(Now(), "dd/mm/yyyy") + "'," + CStr(0) + "," + CStr(datos!adeudo) + "," + CStr(0) + ")"
'            ordensigue = ordensigue + 1
'            datos.MoveNext
'        Wend
'        datos.Close
'        base.Execute "delete from regen"

End Sub
Private Sub gastos()

Dim datos, datos1, datos2 As Recordset
Dim pagos1, gastos, folio As Double
Dim conta As Integer

If txtfecha.Text <> "" Then
        
    'base.Execute "delete from rep_gastos_dia"
        
    conta = 0
    Set datos = Base.OpenRecordset("select * from gastos_dia")
    While Not datos.EOF
        If datos!fecha = CDate(txtfecha.Text) Then
            Base.Execute "insert into regen (cve,cve1,cve2,cve3,cve4,cve5,gasto,fechag,importe,descripcion) values(" + CStr(4) + "," + CStr(0) + "," + CStr(0) + "," + CStr(0) + "," + CStr(4) + "," + CStr(0) + "," + CStr(datos!Gasto) + ",'" + Format(datos!fecha, "dd/mm/yyyy") + "','" + CStr(datos!importe) + "','" + datos!descripcion + "')"
            'base.Execute "insert into regen (cve,gasto,fechag,importe,descripcion) values(" + CStr(4) + "," + CStr(datos!gasto) + ",'" + Format(datos!fecha, "dd/mm/yyyy") + "','" + CStr(datos!importe) + "','" + datos!descripcion + "')"
            conta = 1
        End If
        datos.MoveNext
    Wend
    datos.Close
'    If conta = 1 Then
'        regastos.PrintReport
'    Else
'        MsgBox "No existen gastos registrados para el dia seleccionado", vbInformation, "Reporte de Gastos"
'        txtfecha.SetFocus
'    End If
   
Else
'    MsgBox "El dato de la fecha es incorrecto", vbCritical, "Resumen Diario"
'    txtfecha.Text = Format(Now, "dd/mm/yyyy")
'    txtfecha.SetFocus
End If


End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtfecha.Text = Format(Now, "dd/mm/yyyy")
txtcantidad.Text = 0

Dim suma As Double
Dim datos As Recordset
Dim fecha As Date
fecha = Format(Now, "dd/mm/yyyy")
Set datos = Base.OpenRecordset("Select * from Sobrantes where cdbl(fecha) = " + CStr(CDbl(fecha)))
suma = 0
While Not datos.EOF
    If CStr(datos!tipo) = "S" Then
        suma = suma + CDbl(datos!cantidad)
    Else
        If CStr(datos!tipo) = "F" Then
            suma = suma - CDbl(datos!cantidad)
        End If
    End If
datos.MoveNext
Wend

If suma < 0 Then
    optfaltante.Value = True
    suma = suma - (2 * suma)
Else
    If suma >= 0 Then
        optsobrante.Value = True
    End If
End If

txtcantidad.Text = suma

End Sub

