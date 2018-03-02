VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form periodofrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Periodo"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   870
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   1230
      TabIndex        =   2
      Top             =   870
      Width           =   1335
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtFechaInicial 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1845
      _Version        =   65537
      _ExtentX        =   3254
      _ExtentY        =   609
      _StockProps     =   93
      ShowCentury     =   -1  'True
      Mask            =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtFechaFinal 
      Height          =   345
      Left            =   2130
      TabIndex        =   1
      Top             =   360
      Width           =   1845
      _Version        =   65537
      _ExtentX        =   3254
      _ExtentY        =   609
      _StockProps     =   93
      ShowCentury     =   -1  'True
      Mask            =   2
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Final"
      Height          =   225
      Left            =   2160
      TabIndex        =   5
      Top             =   90
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicial"
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   90
      Width           =   1815
   End
End
Attribute VB_Name = "periodofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strFechaInicial As String
Public strFechaFinal As String

Private Sub cmdAceptar_Click()

    If DateDiff("d", dtFechaInicial, dtFechaFinal) < 0 Then
        MsgBox "¡Revise el periodo, la fecha final debe ser posterior a la fecha inicial!", vbInformation + vbOKOnly
    End If
    
    strFechaInicial = dtFechaInicial.Text
    strFechaFinal = dtFechaFinal.Text
    
    Unload Me
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub
