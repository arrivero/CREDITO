VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form dlgMsgBox 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDITOS"
   ClientHeight    =   2880
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel lblMensaje 
      Height          =   825
      Left            =   2190
      TabIndex        =   1
      Top             =   1170
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1455
      _Version        =   196609
      Font3D          =   1
      MarqueeStyle    =   3
      ForeColor       =   16711680
      BackColor       =   -2147483634
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodShowPct    =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Height          =   2685
      Left            =   90
      Picture         =   "dlgMsgBox.frx":0000
      ScaleHeight     =   2625
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   90
      Width           =   1935
   End
End
Attribute VB_Name = "dlgMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public iResult As Integer
Public strMensaje As String

Private Sub Form_Load()

    lblMensaje = strMensaje
        
End Sub

