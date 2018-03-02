VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form convierteValorCadenafrm 
   Caption         =   "Form1"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin EditLib.fpCurrency fpMonto 
      Height          =   435
      Left            =   1410
      TabIndex        =   1
      Top             =   480
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   767
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ","
      UseSeparator    =   -1  'True
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir"
      Height          =   465
      Left            =   3630
      TabIndex        =   0
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label lblMontoLetra 
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   1170
      Width           =   7665
   End
End
Attribute VB_Name = "convierteValorCadenafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrUnidad(0 To 19) As String

Private arrDecena(0 To 9) As String

Private arrCentena(1 To 9) As String

Private Sub Form_Load()

    InicializarArrays
    
End Sub

Private Sub InicializarArrays()
    
    'Asignar los valores
    arrUnidad(1) = ""
    arrUnidad(1) = "UN"
    arrUnidad(2) = "DOS"
    arrUnidad(3) = "TRES"
    arrUnidad(4) = "CUATRO"
    arrUnidad(5) = "CINCO"
    arrUnidad(6) = "SEIS"
    arrUnidad(7) = "SIETE"
    arrUnidad(8) = "OCHO"
    arrUnidad(9) = "NUEVE"
    arrUnidad(10) = "DIEZ"
    arrUnidad(11) = "ONCE"
    arrUnidad(12) = "DOCE"
    arrUnidad(13) = "TRECE"
    arrUnidad(14) = "CATORCE"
    arrUnidad(15) = "QUINCE"
    arrUnidad(16) = "DIEZ Y SEIS"
    arrUnidad(17) = "DIEZ Y SIETE"
    arrUnidad(18) = "DIEZ Y OCHO"
    arrUnidad(19) = "DIEZ Y NUEVE"
    
    '
    arrDecena(0) = ""
    arrDecena(1) = "DIEZ"
    arrDecena(2) = "VEINTE"
    arrDecena(3) = "TREINTA"
    arrDecena(4) = "CUARENTA"
    arrDecena(5) = "CINCUENTA"
    arrDecena(6) = "SESENTA"
    arrDecena(7) = "SETENTA"
    arrDecena(8) = "OCHENTA"
    arrDecena(9) = "NOVENTA"
    '
    arrCentena(1) = "CIEN"
    arrCentena(2) = "DOSCIENTOS"
    arrCentena(3) = "TRESCIENTOS"
    arrCentena(4) = "CUATROCIENTOS"
    arrCentena(5) = "QUINIENTOS"
    arrCentena(6) = "SEISCIENTOS"
    arrCentena(7) = "SETECIENTOS"
    arrCentena(8) = "OCHOCIENTOS"
    arrCentena(9) = "NOVECIENTOS"

End Sub

Private Sub Command1_Click()
    
    lblMontoLetra.Caption = convierteMontoConLentra(fnstrValor(fpMonto))

End Sub

Private Function convierteMontoConLentra(strNo As String) As String
    
    Dim strNoEntero, strDecimal As String
    Dim strSubEntero As String
    Dim strMillon As String
    Dim strCantidadConLetra, strMontoConLetra, strCantidadConLetraTemp As String
    Dim iPos, iBloque As Integer
    Dim bMiles, bUnidades As Boolean
    
    iPos = InStr(1, strNo, ".")
    
    strNoEntero = Mid(strNo, 1, iPos - 1)
    strDecimal = Mid(strNo, iPos + 1, 2)
    
    iBloque = 1
    
    Do
        strSubEntero = Right(strNoEntero, 3)
            
        strCantidadConLetraTemp = obtenCantidadLetra(strSubEntero)
                
        Select Case iBloque
            Case Is = 1 'cienes
                strCantidadConLetra = strCantidadConLetraTemp
                If Trim(strCantidadConLetraTemp) <> "" Then
                    bUnidades = True
                Else
                    bUnidades = False
                End If
            Case Is = 2 'miles
                If Trim(strCantidadConLetraTemp) <> "" Then
                    strCantidadConLetra = strCantidadConLetraTemp & " MIL " & strCantidadConLetra
                    bMiles = True
                Else
                    bMiles = False
                End If
            Case Is = 3 'millones
                If Trim(strCantidadConLetraTemp) <> "" Then
                    If Trim(strCantidadConLetraTemp) = "UN" Then
                        If bUnidades = True Then
                            strMillon = " MILLON "
                        Else
                            If bMiles = True Then
                                strMillon = " MILLON "
                            Else
                                strMillon = " MILLON DE "
                            End If
                        End If
                    Else
                        If bUnidades = True Then
                            strMillon = " MILLONES "
                        Else
                            If bMiles = True Then
                                strMillon = " MILLONES "
                            Else
                                strMillon = " MILLONES DE "
                            End If
                        End If
                    End If
                    strCantidadConLetra = strCantidadConLetraTemp & strMillon & strCantidadConLetra
                End If
            Case Is = 4 'miles de millones
                strCantidadConLetra = strCantidadConLetraTemp & " MIL " & strCantidadConLetra
            Case Is = 5 'billones
                strCantidadConLetra = strCantidadConLetraTemp & " BILLONES " & strCantidadConLetra
        End Select
        
        iBloque = iBloque + 1
        
        If Len(strNoEntero) - 3 > 0 Then
            strNoEntero = Mid(strNoEntero, 1, Len(strNoEntero) - 3)
        Else
            Exit Do
        End If
        
    Loop While Len(strNoEntero) > 0
    
    If strCantidadConLetra <> "" Then
        Dim strPesos As String
        If strCantidadConLetra = "UN" Then
            strPesos = " PESO "
        Else
            strPesos = " PESOS "
        End If
        
        convierteMontoConLentra = "(" & strCantidadConLetra & strPesos & strDecimal & "/100 M.N.)"
    Else
        convierteMontoConLentra = "(CERO PESOS" & strDecimal & "/100 M.N.)"
    End If

End Function

Private Function obtenCantidadLetra(strSubEntero As String) As String
    
    Dim strSubSubEntero, strCentena, strDecena, strUnidad As String
    Dim ivalorCentena, ivalorDecena, ivalorUnidad As Integer
    
    If Val(strSubEntero) > 99 Then
        ivalorCentena = Val(Left(strSubEntero, 1))
        If ivalorCentena = 1 Then
            ivalorDecena = Val(Right(strSubEntero, 2))
            If ivalorDecena > 0 Then
                strCentena = "CIENTO"
            Else
                ivalorCentena = Val(Left(strSubEntero, 1))
                strCentena = arrCentena(ivalorCentena)
            End If
        Else
            ivalorCentena = Val(Left(strSubEntero, 1))
            strCentena = arrCentena(ivalorCentena)
        End If
        
        strSubSubEntero = Right(strSubEntero, 2)
        
        If Val(strSubSubEntero) < 20 Then
            strUnidad = arrUnidad(Val(strSubSubEntero))
        Else
            strDecena = arrDecena(Val(Left(strSubSubEntero, 1)))
            If Val(Right(strSubSubEntero, 1)) > 0 Then
                strUnidad = "Y "
                strUnidad = strUnidad & arrUnidad(Val(Right(strSubSubEntero, 1)))
            Else
                strUnidad = arrUnidad(Val(Right(strSubSubEntero, 1)))
            End If
        End If
        
    Else
    
        If Val(strSubEntero) < 20 Then
            strUnidad = arrUnidad(Val(strSubEntero))
        Else
            ivalorDecena = Val(Left(strSubEntero, 1))
            ivalorUnidad = Val(Right(strSubEntero, 1))
            strDecena = arrDecena(ivalorDecena)
            If ivalorUnidad > 0 Then
                strUnidad = "Y "
                strUnidad = strUnidad & arrUnidad(ivalorUnidad)
            Else
                strUnidad = arrUnidad(ivalorUnidad)
            End If
        End If
        
    End If
    
    If strCentena <> "" Then
        If strDecena <> "" Then
            If strUnidad <> "" Then
                obtenCantidadLetra = strCentena & " " & strDecena & " " & strUnidad
            Else
                obtenCantidadLetra = strCentena & " " & strDecena
            End If
        Else
            If strUnidad <> "" Then
                obtenCantidadLetra = strCentena & " " & strUnidad
            Else
                obtenCantidadLetra = strCentena
            End If
        End If
    Else
    
        If strDecena <> "" Then
            If strUnidad <> "" Then
                obtenCantidadLetra = strDecena & " " & strUnidad
            Else
                obtenCantidadLetra = strDecena
            End If
        Else
            If strUnidad <> "" Then
                obtenCantidadLetra = strUnidad
            Else
                obtenCantidadLetra = ""
            End If
        End If
    
    End If
    
End Function

