VERSION 5.00
Begin VB.Form gastosTrabajoInternosfrm 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Gastos"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Aceptar"
      Height          =   465
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1740
      Width           =   1125
   End
   Begin VB.OptionButton opAutomaticoManual 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Trabajo"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   1170
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.OptionButton opAutomaticoManual 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Internos"
      Height          =   255
      Index           =   1
      Left            =   2370
      TabIndex        =   0
      Top             =   1170
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "-Trabajo, para captura de gastos de trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "-Internos, para registro de gastos operativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "gastosTrabajoInternosfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iTrabajoInternos As Integer

Private Sub cmdAceptar_Click()
    
    If opAutomaticoManual.Item(0).Value = False And opAutomaticoManual.Item(1).Value = False Then
        MsgBox "¡Seleccione modo de registro de gastos.!", vbInformation + vbOKOnly
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    iTrabajoInternos = 1
End Sub

Private Sub opAutomaticoManual_Click(Index As Integer)
    
    If opAutomaticoManual.Item(0).Value = True Then
        iTrabajoInternos = 1
    Else
        iTrabajoInternos = 0
    End If
    
End Sub

