VERSION 5.00
Object = "{E2585150-2883-11D2-B1DA-00104B9E0750}#3.0#0"; "ssresz30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form configuracionfrm 
   BorderStyle     =   0  'None
   Caption         =   "Configuración"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel pnlConfiguracion 
      Height          =   3060
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   230
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   5398
      _Version        =   196608
      BevelWidth      =   0
      BorderWidth     =   0
      AutoSize        =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
      Begin VB.TextBox txtPswReConfirmar 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   25
         Text            =   "abcdefg"
         Top             =   1800
         Width           =   1860
      End
      Begin VB.TextBox txtPswConfirmar 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   24
         Text            =   "abcdefg"
         Top             =   1440
         Width           =   1860
      End
      Begin VB.TextBox txtpwd 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   23
         Text            =   "abcdefg"
         Top             =   1080
         Width           =   1860
      End
      Begin VB.CommandButton cmdAlta 
         Caption         =   "Alta"
         Height          =   495
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2370
         Width           =   1215
      End
      Begin EditLib.fpText txtpwdold 
         Height          =   285
         Left            =   2010
         TabIndex        =   5
         Top             =   1080
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
         _ExtentY        =   503
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "fpText1"
         CharValidationText=   ""
         MaxLength       =   10
         MultiLine       =   0   'False
         PasswordChar    =   "*"
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin Threed.SSCheck ckbTipo 
         Height          =   315
         Left            =   4050
         TabIndex        =   8
         Top             =   570
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   196608
         Caption         =   "Gerente"
      End
      Begin VB.ComboBox cmbUsuario 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2010
         TabIndex        =   4
         Text            =   "cmbUsuario"
         Top             =   570
         Width           =   1815
      End
      Begin EditLib.fpText txtPswConfirmarold 
         Height          =   285
         Left            =   2010
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
         _ExtentY        =   503
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "fpText1"
         CharValidationText=   ""
         MaxLength       =   10
         MultiLine       =   0   'False
         PasswordChar    =   "*"
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText txtPswReConfirmarold 
         Height          =   285
         Left            =   2010
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
         _Version        =   196608
         _ExtentX        =   3201
         _ExtentY        =   503
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "fpText1"
         CharValidationText=   ""
         MaxLength       =   10
         MultiLine       =   0   'False
         PasswordChar    =   "*"
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label12 
         Caption         =   "* Requerido para alta y/o cambio de password"
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   4020
         TabIndex        =   21
         Top             =   1020
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "** Requerido para cambio de password"
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   4050
         TabIndex        =   20
         Top             =   1770
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "** Reconfirma Password:"
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label9 
         Caption         =   "* Confirma Password:"
         Height          =   255
         Left            =   330
         TabIndex        =   18
         Top             =   1470
         Width           =   1635
      End
      Begin VB.Label Label8 
         Caption         =   "Password:"
         Height          =   255
         Left            =   990
         TabIndex        =   11
         Top             =   1095
         Width           =   945
      End
      Begin VB.Label Label7 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   615
         Width           =   765
      End
   End
   Begin Threed.SSPanel pnlConfiguracion 
      Height          =   3060
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   230
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5398
      _Version        =   196608
      BevelWidth      =   0
      BorderWidth     =   0
      AutoSize        =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
      Begin MSMask.MaskEdBox txtcuatro 
         Height          =   285
         Left            =   3435
         TabIndex        =   37
         Top             =   1665
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   1
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txttres 
         Height          =   285
         Left            =   3435
         TabIndex        =   36
         Top             =   1305
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   1
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtdos 
         Height          =   285
         Left            =   3435
         TabIndex        =   35
         Top             =   945
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   1
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtuno 
         Height          =   285
         Left            =   3435
         TabIndex        =   34
         Top             =   585
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   1
         Mask            =   "#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtcuarto 
         Height          =   285
         Left            =   2160
         TabIndex        =   33
         Top             =   1650
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txttercero 
         Height          =   285
         Left            =   2160
         TabIndex        =   32
         Top             =   1290
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtsegundo 
         Height          =   285
         Left            =   2160
         TabIndex        =   31
         Top             =   930
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtprimero 
         Height          =   285
         Left            =   2160
         TabIndex        =   30
         Top             =   570
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdactualizar 
         Caption         =   "Registra Cambios"
         Height          =   465
         Left            =   2550
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "1er atraso:"
         Height          =   255
         Left            =   960
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Días atrasados"
         Height          =   195
         Left            =   2160
         TabIndex        =   16
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Días a cobrar"
         Height          =   255
         Left            =   3450
         TabIndex        =   15
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "2ndo atraso:"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "3er atraso:"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "4to atraso:"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
   End
   Begin Threed.SSPanel pnlConfiguracion 
      Height          =   3060
      Index           =   2
      Left            =   0
      TabIndex        =   26
      Top             =   230
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5398
      _Version        =   196608
      BevelWidth      =   0
      BorderWidth     =   0
      AutoSize        =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   0   'False
      Begin VB.CommandButton cmdAltaEmpleado 
         Caption         =   "Alta"
         Height          =   330
         Left            =   4770
         TabIndex        =   28
         Top             =   2475
         Width           =   1095
      End
      Begin VB.CommandButton cmdActualizaEmpleado 
         Caption         =   "Actualizar"
         Height          =   330
         Left            =   3330
         TabIndex        =   27
         Top             =   2475
         Width           =   1095
      End
      Begin FPSpread.vaSpread sprEmpleado 
         Height          =   2385
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   5880
         _Version        =   196608
         _ExtentX        =   10372
         _ExtentY        =   4207
         _StockProps     =   64
         BorderStyle     =   0
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
         MaxCols         =   4
         MaxRows         =   20
         OperationMode   =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "configuracionfrm.frx":0000
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Salir"
      Height          =   495
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2610
      Width           =   1215
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   1110
      Top             =   3030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   196608
      MinFontSize     =   1
      MaxFontSize     =   100
   End
   Begin ComctlLib.TabStrip tabConfiguracion 
      Height          =   3195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   5636
      TabWidthStyle   =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Usuarios"
            Key             =   "CONUSU"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reglas Moratorios"
            Key             =   "CONREGMOR"
            Object.Tag             =   ""
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1980
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "configuracionfrm.frx":046E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "configuracionfrm.frx":0648
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "configuracionfrm.frx":0822
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "configuracionfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miFrameActivo As Integer

Private bCambio As Boolean
Private bAlta As Boolean

Private miEstadoCivil As Integer
Private miEstadoEmpleado As Integer

Private Sub cmbUsuario_Change()
    bAlta = True
End Sub

Private Sub cmbUsuario_Click()

    Dim Registro As New Collection
    Dim oCampo As New Campo
    
    Dim oUsuario As New Usuario
    Dim iUsuario As Integer
    iUsuario = cmbUsuario.ItemData(cmbUsuario.ListIndex)
    
    oUsuario.busca iUsuario
    
    If oUsuario.cDatos.Count > 0 Then
    
        Set Registro = oUsuario.cDatos(1)
        
        'Set oCampo = Registro(2) 'Password
        txtpwd = ""
        txtPswConfirmar = ""
        txtPswReConfirmar = ""
        
        Set oCampo = Registro(3) 'Pefil
        
        If oCampo.Valor = 1 Then
            ckbTipo.Value = ssCBChecked
        Else
            ckbTipo.Value = ssCBUnchecked
        End If
        
        'Set oCampo = Registro(4) 'Estado
        cmdAlta.Caption = "Cambio"
        
        Label10.Visible = True
        Label11.Visible = True
        Me.txtPswReConfirmar.Visible = True
        
    End If
            
    Set oUsuario = Nothing
    
End Sub

Private Sub cmdAlta_Click()

    Dim oUsuario As New Usuario
    Dim Registro As New Collection
    Dim oCampo As New Campo
    Dim iTipoUsuario As Integer
    Dim strUsuario As String
        
    'Valida que el Usuario no sea mayor a 10 Caracteres
    If Len(cmbUsuario.Text) > 10 Then
        MsgBox "La longitud máxisma de un Usuario es 10, verifique por favor!", vbInformation + vbOKOnly
        cmbUsuario.SetFocus
        Exit Sub
    End If
    
    If Len(txtpwd.Text) <= 0 Then
        MsgBox "Debe ingresar el password actual, por favor verifique!", vbInformation + vbOKOnly
        txtpwd.SetFocus
        Exit Sub
    End If
    
    If Len(txtPswConfirmar.Text) <= 0 Then
        MsgBox "Debe ingresar el password nuevo, por favor verifique!", vbInformation + vbOKOnly
        txtPswConfirmar.SetFocus
        Exit Sub
    End If
    
        
    If cmdAlta.Caption = "Alta" Then
        
        'Valida que el passwor es igual a su confirmación
        If txtpwd.Text <> txtPswConfirmar Then
            MsgBox "La confirmación de password es diferente, por favor verifique!", vbInformation + vbOKOnly
            txtpwd.SetFocus
            Exit Sub
        End If
        
        'Valida que el Usuario no existe ya en la lista de usuarios
        oUsuario.buscaUsuario cmbUsuario.Text
        
        If oUsuario.bDatos = True Then
        
            MsgBox "El usuario " + cmbUsuario.Text + " ya existe, por favor verifique!", vbInformation + vbOKOnly
            cmbUsuario.SetFocus
            Exit Sub
        
        End If
        
        'Registra el nuevo usuario
        If ckbTipo.Value = -1 Then
            iTipoUsuario = USUARIO_GERENTE
        Else
            iTipoUsuario = USUARIO_USUARIO
        End If
        
        oUsuario.registra iTipoUsuario, cmbUsuario.Text, txtpwd.Text
        
        'Actualiza la lista de usuarios
        strUsuario = cmbUsuario.Text
        If oUsuario.catalogoUsuarios Then
            fnLlenaComboCollecion cmbUsuario, oUsuario.cDatos, 0, ""
            
            'Busca el nuevo usuario y haslo activo en el combo
            oUsuario.buscaUsuario strUsuario
        
            If oUsuario.bDatos = True Then
                
                Set Registro = oUsuario.cDatos(1)
                Set oCampo = Registro(1)
                
                fnBuscaElemento cmbUsuario, oCampo.Valor
                
            End If
            
        End If
            
        MsgBox "El usuario " + cmbUsuario.Text + " fué registrado con éxito!", vbInformation + vbOKOnly
        
        txtpwd = ""
        txtPswConfirmar = ""
        txtPswReConfirmar = ""
        
    Else
        Dim iClave As Integer
        iClave = 0
        If Len(txtPswConfirmar.Text) > 0 Or Len(txtPswReConfirmar.Text) > 0 Then
            
            'Valida que el password actual del Usuario
            oUsuario.busca cmbUsuario.ItemData(cmbUsuario.ListIndex)
            If oUsuario.cDatos.Count > 0 Then
                
                Dim strPassword As String
                
                Set Registro = oUsuario.cDatos(1)
                
                Set oCampo = Registro(2) 'Password
                
                strPassword = oCampo.Valor
                
                If strPassword <> txtpwd Then
                    MsgBox "Su password no es correcto, verifique por favor!", vbInformation + vbOKOnly
                    txtpwd.SetFocus
                    Exit Sub
                End If
                            
            End If
        
            'Verifica que el password nuevo y el de confirmación son iguales
            If txtPswConfirmar.Text <> txtPswReConfirmar.Text Then
                MsgBox "La confirmación de password es diferente, por favor verifique!", vbInformation + vbOKOnly
                txtPswReConfirmar.SetFocus
                Exit Sub
            End If
            
            iClave = 1
            
        
            'actualiza el usuario
            If ckbTipo.Value = -1 Then
                iTipoUsuario = USUARIO_GERENTE
            Else
                iTipoUsuario = USUARIO_USUARIO
            End If
            
            oUsuario.actualizaClave cmbUsuario.ItemData(cmbUsuario.ListIndex), txtPswConfirmar.Text, iTipoUsuario, iClave
            
            MsgBox "El usuario " + cmbUsuario.Text + " fué actualizado con éxito!", vbInformation + vbOKOnly
            
            txtpwd = ""
            txtPswConfirmar = ""
            txtPswReConfirmar = ""
        
        End If
        
    End If
    
    Set oUsuario = Nothing
    
End Sub


Private Sub Form_Load()

    Dim cMoratorios As New Collection
    Dim cRegistro As New Collection
    Dim oCampo As New Campo
    
    Dim oCredito As New credito
    oCredito.obtenMoratorios
    
    bCambio = False
    
    Set cMoratorios = oCredito.cDatos

    Set cRegistro = cMoratorios(1)
    Set oCampo = cRegistro(1)
    txtprimero.Tag = oCampo.Valor
    Set oCampo = cRegistro(2)
    txtprimero.Text = oCampo.Valor
    Set oCampo = cRegistro(3)
    txtuno.Text = oCampo.Valor

    Set cRegistro = cMoratorios(2)
    Set oCampo = cRegistro(1)
    txtsegundo.Tag = oCampo.Valor
    Set oCampo = cRegistro(2)
    txtsegundo.Text = oCampo.Valor
    Set oCampo = cRegistro(3)
    txtdos.Text = oCampo.Valor

    Set cRegistro = cMoratorios(3)
    Set oCampo = cRegistro(1)
    txttercero.Tag = oCampo.Valor
    Set oCampo = cRegistro(2)
    txttercero.Text = oCampo.Valor
    Set oCampo = cRegistro(3)
    txttres.Text = oCampo.Valor

    Set cRegistro = cMoratorios(4)
    Set oCampo = cRegistro(1)
    txtcuarto.Tag = oCampo.Valor
    Set oCampo = cRegistro(2)
    txtcuarto.Text = oCampo.Valor
    Set oCampo = cRegistro(3)
    txtcuatro.Text = oCampo.Valor
    Set oCredito = Nothing
    
    'Carga el catalogo de cobradores
    Dim oUsuario As New Usuario
    If oUsuario.catalogoUsuarios Then
        fnLlenaComboCollecion cmbUsuario, oUsuario.cDatos, 0, ""
        cmbUsuario.ListIndex = -1
    End If
    Set oUsuario = Nothing
    
    If giTipoUsuario <> USUARIO_GERENTE Then

        fnBuscaTextoCombo cmbUsuario, gstrUsuario
        Me.ckbTipo.Visible = False
        tabConfiguracion.Enabled = False
        cmbUsuario.Enabled = False
        
    End If
    
    Dim oEmpleado As New Empleado
    
        If oEmpleado.obtenLista() = True Then
        
            Call fnLlenaTablaCollection(sprEmpleado, oEmpleado.cDatos)
            
        End If
        
    Set oEmpleado = Nothing
    
End Sub



Private Sub tabConfiguracion_Click()

   pnlConfiguracion(obtenFrame(tabConfiguracion.SelectedItem.key)).ZOrder 0

'    If giTipoUsuario = USUARIO_GERENTE Then
'
'        If tabConfiguracion.SelectedItem.Index - 1 = miFrameActivo Then Exit Sub ' No need to change frame.
'
'        ' Comosea, oculta el frame anterior, muestra el nuevo.
'        pnlConfiguracion(tabConfiguracion.SelectedItem.Index - 1).Visible = True
'        pnlConfiguracion(miFrameActivo).Visible = False
'
'        miFrameActivo = tabConfiguracion.SelectedItem.Index - 1
'
'    End If
'
End Sub

Private Function obtenFrame(key As String) As Integer
    
    Dim iFrame As Integer
    Select Case key
        Case Is = "CONUSU"
            iFrame = 0
        Case Is = "CONREGMOR"
            iFrame = 1
        Case Is = "CONPERSONAL"
            iFrame = 2
    End Select

    obtenFrame = iFrame
    
End Function

Private Sub cmdActualizar_Click()

    If bCambio = True Then
    
        Dim Registros As New Collection
        Dim Registro As Collection
        Dim oCampo As New Campo
        
        Set Registro = New Collection
        Registro.Add oCampo.CreaCampo(adInteger, , , 1)
        Registro.Add oCampo.CreaCampo(adInteger, , , txtprimero.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , txtuno.Text)
        Registros.Add Registro
        
        Set Registro = New Collection
        Registro.Add oCampo.CreaCampo(adInteger, , , 2)
        Registro.Add oCampo.CreaCampo(adInteger, , , txtsegundo.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , txtdos.Text)
        Registros.Add Registro

        Set Registro = New Collection
        Registro.Add oCampo.CreaCampo(adInteger, , , 3)
        Registro.Add oCampo.CreaCampo(adInteger, , , txttercero.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , txttres.Text)
        Registros.Add Registro

        Set Registro = New Collection
        Registro.Add oCampo.CreaCampo(adInteger, , , 4)
        Registro.Add oCampo.CreaCampo(adInteger, , , txtcuarto.Text)
        Registro.Add oCampo.CreaCampo(adInteger, , , txtcuatro.Text)
        Registros.Add Registro

        Dim oCredito As New credito
        oCredito.actualizaMoratorios Registros
        Set oCredito = Nothing
        
        MsgBox "Las reglas de Moratorios, han sido actualizadas", vbInformation + vbOKOnly
        
        cmdactualizar.Enabled = False
        
     End If
     
End Sub

Private Sub txtcuarto_Change()
    bCambio = True
End Sub

Private Sub txtcuatro_Change()
    bCambio = True
End Sub

Private Sub txtdos_Change()
    bCambio = True
End Sub

Private Sub txtprimero_Change()
    bCambio = True
End Sub

Private Sub txtpwd_GotFocus()
    
    If cmbUsuario.ListIndex = -1 Then
        
        txtpwd.Text = ""
        cmdAlta.Caption = "Alta"
        
        Label10.Visible = False
        Label11.Visible = False
        Me.txtPswReConfirmar.Visible = False
        
        Me.txtPswConfirmar.Text = ""
        Me.ckbTipo.Value = ssCBUnchecked
        
    End If
    
End Sub

Private Sub txtsegundo_Change()
    bCambio = True
End Sub

Private Sub txttercero_Change()
    bCambio = True
End Sub

Private Sub txttres_Change()
    bCambio = True
End Sub

Private Sub txtuno_Change()
    bCambio = True
End Sub

Private Sub cmdsalir_Click()
    sicPrincipalfrm.Caption = ""
    sicPrincipalfrm.pnlTitulo.Caption = "Solución Integral de Administración de Creditos"
    despliegaVentana portadafrm, WND_PORTADA
End Sub

'ADMINISTRACIÓN DE EMPLEADOS

Private Sub cmdAltaEmpleado_Click()
    
    empleadoAlta.Show vbModal
    
    If empleadoAlta.bAlta = True Then
    
        actualizaListaEmpleados
    
    End If
    
End Sub

Private Sub cmdActualizaEmpleado_Click()
    
    sprEmpleado.Row = sprEmpleado.ActiveRow
    sprEmpleado.Col = 4
        
    empleadoCambio.iEmpleado = Val(sprEmpleado.Text)
    empleadoCambio.Show vbModal
    
    If empleadoCambio.bAlta = True Then
    
        actualizaListaEmpleados
        
    End If
    
End Sub

Private Function actualizaListaEmpleados()

    Dim oEmpleado As New Empleado
    
        If oEmpleado.obtenLista() = True Then
        
            Call fnLlenaTablaCollection(sprEmpleado, oEmpleado.cDatos)
            
        End If
        
    Set oEmpleado = Nothing

End Function
