VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form automaticoManualfrm 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modo de registro de pagos"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   1020
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1138
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   -2147483645
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Modo"
      Begin VB.OptionButton opAutomaticoManual 
         BackColor       =   &H80000003&
         Caption         =   "Manual"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2490
         TabIndex        =   4
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton opAutomaticoManual 
         BackColor       =   &H80000003&
         Caption         =   "Automático"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H80000003&
      Caption         =   "Aceptar"
      Height          =   465
      Left            =   1440
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1740
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000003&
      Caption         =   "-Selección Automática, para registro automático de pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   30
      TabIndex        =   5
      Top             =   390
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "-Selección Manual, para captura de pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   3975
   End
End
Attribute VB_Name = "automaticoManualfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iAutomaticoManual As Integer

Private Sub cmdAceptar_Click()
    
    If opAutomaticoManual.Item(0).Value = False And opAutomaticoManual.Item(1).Value = False Then
        MsgBox "¡Seleccione modo de registro de pagos.!", vbInformation + vbOKOnly
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    iAutomaticoManual = 1
End Sub

Private Sub opAutomaticoManual_Click(Index As Integer)
    
    If opAutomaticoManual.Item(0).Value = True Then
        iAutomaticoManual = 1
    Else
        iAutomaticoManual = 0
    End If
    
End Sub
