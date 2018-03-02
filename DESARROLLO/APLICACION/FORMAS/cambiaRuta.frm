VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form cambiaRuta 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Intercambio de Ruta"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H8000000E&
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1290
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H8000000E&
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1290
      Width           =   1335
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1035
      Left            =   30
      TabIndex        =   0
      Top             =   210
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   1826
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4230
         TabIndex        =   3
         Top             =   510
         Width           =   2175
      End
      Begin VB.ComboBox cbCobrador 
         Height          =   315
         Left            =   990
         TabIndex        =   1
         Top             =   510
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "que cubre:"
         Height          =   255
         Left            =   3270
         TabIndex        =   4
         Top             =   570
         Width           =   915
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "A cubrir:"
         Height          =   255
         Left            =   270
         TabIndex        =   2
         Top             =   540
         Width           =   825
      End
   End
End
Attribute VB_Name = "cambiaRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
